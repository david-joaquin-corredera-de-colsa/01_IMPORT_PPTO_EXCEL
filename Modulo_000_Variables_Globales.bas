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

' Constante para version de la Macro
Public Const CONST_MACRO_VERSION As String = "Macro - Version 20250616 - 10:06:01"

' Constante para Scenario Admitido
Public Const CONST_ESCENARIO_ADMITIDO As String = "BUDGET_OS"

' Constante para Ultimo Mes De Carga
Public Const CONST_ULTIMO_MES_DE_CARGA As String = "M12"

' Constante numero de hojas historicas visibles
Public Const CONST_NUM_HOJAS_HCAS_VISIBLES_ENVIO As Integer = 5

' Constantes para el borrado de hojas antiguas y no deseadas
Public Const CONST_NUM_HOJAS_HCAS_IMPORT_TARGET As Integer = CONST_NUM_HOJAS_HCAS_VISIBLES_ENVIO      'Este numero normalmente lo fijamos a CONST_NUM_HOJAS_HCAS_VISIBLES_ENVIO = 5
Public Const CONST_BORRAR_OTRAS_HOJAS_PREFIJO_00 As Boolean = True
Public Const CONST_BORRAR_OTRAS_HOJAS_PREFIJO_IMPORT As Boolean = True
Public Const CONST_BORRAR_HOJAS_SIN_PREFIJOS_00_IMPORT As Boolean = True

' Constante para columna de la Entity en la hoja de datos importados/envio/comprobacion
Public Const CONST_COLUMNA_ENTITY As Integer = 4

' Constantes para mostrar o no mensajes durante la ejecución
Public Const CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS As Boolean = False

' Constantes para control de errores
Public Const ERROR_BASE_IMPORT As Long = vbObjectError + 1000

' Constantes para nombres de HOJAS TECNICAS
Public Const CONST_HOJA_EJECUTAR_PROCESOS As String = "00_Ejecutar_Procesos"
Public Const CONST_HOJA_INVENTARIO As String = "01_Inventario"
Public Const CONST_HOJA_LOG As String = "02_Log"
Public Const CONST_HOJA_USERNAME As String = "05_Username"
Public Const CONST_HOJA_DELIMITADORES_ORIGINALES As String = "06_Delimitadores_Originales"
Public Const CONST_HOJA_REPORT_PL As String = "09_Report_PL"
Public Const CONST_HOJA_REPORT_PL_AH As String = "10_Report_PL_AH"

' Constantes para elegir si cada HOJAS TECNICA debe quedar Hidden - valores posibles xlSheetVisible = -1 | xlSheetHidden = 0 | xlSheetVeryHidden = 2
Public Const CONST_HOJA_EJECUTAR_PROCESOS_VISIBLE As Integer = xlSheetVisible
Public Const CONST_HOJA_INVENTARIO_VISIBLE As Integer = xlSheetVisible
Public Const CONST_HOJA_LOG_VISIBLE As Integer = xlSheetVisible
Public Const CONST_HOJA_USERNAME_VISIBLE As Integer = xlSheetHidden
Public Const CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE As Integer = xlSheetHidden
Public Const CONST_HOJA_REPORT_PL_VISIBLE As Integer = xlSheetVisible
Public Const CONST_HOJA_REPORT_PL_AH_VISIBLE As Integer = xlSheetHidden

' Constantes para prefijos en los nombres de las HOJAS TECNICAS
Public Const CONST_PREFIJO_HOJA_IMPORTACION As String = "Import_"
Public Const CONST_PREFIJO_HOJA_IMPORTACION_WORKING As String = "Import_Working_"
Public Const CONST_PREFIJO_HOJA_IMPORTACION_ENVIO As String = "Import_Envio_"
Public Const CONST_PREFIJO_HOJA_IMPORTACION_COMPROBACION As String = "Import_Comprob_"
Public Const CONST_PREFIJO_HOJA_X_BORRAR_ENVIO_PREVIO As String = "Del_Prev_Envio_"

' Constante para prefijo hoja de backup
Public Const CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO As String = "BK_"



' ============================================================================================
' CONSTANTES GLOBALES PARA POSICION DIMENSIONES A "MODIFICAR" EN HOJA DEL INFORME PL ADHOC
' ============================================================================================
' Constantes para filas:
Public Const CONST_FILA_SCENARIO As Integer = 1
Public Const CONST_FILA_YEAR As Integer = 2
Public Const CONST_FILA_PERIOD As Integer = 3
Public Const CONST_FILA_VIEW As Integer = 4
Public Const CONST_FILA_ENTITY As Integer = 5
Public Const CONST_FILA_VALUE As Integer = 6
Public Const CONST_FILA_ICP As Integer = 7
Public Const CONST_FILA_C4 As Integer = 8
Public Const CONST_FILA_C2_ACTIVITY As Integer = 9
Public Const CONST_FILA_C3_BUSINESS As Integer = 10
' Constantes para COLUMNAS:
Public Const CONST_COLUMNA_INICIAL_HEADERS As Integer = 3
Public Const CONST_COLUMNA_FINAL_HEADERS As Integer = 15
Public Const CONST_COLUMNA_ADICIONAL_HEADERS As Integer = 16


' Constantes para designar las celdas clave de la hoja de Username/Password
Public Const CONST_CELDA_USERNAME As String = "C2"
Public Const CONST_CELDA_HEADER_USERNAME As String = "B2" 'Contendra la string "Username:"
Public Const CONST_VALOR_HEADER_USERNAME As String = "Username:" 'Valor que se asigna a la celda anterior

' Constantes para etiquetas de procesamiento de líneas (NUEVAS) en las hojas de importacion/working/comprobacion/envio
Public Const CONST_TAG_LINEA_TRATADA As String = "Linea_Tratada"
Public Const CONST_TAG_LINEA_SUMA As String = "Linea_Suma"
Public Const CONST_TAG_LINEA_REPETIDA As String = "Linea_Repetida"

' ============================================================================================
' CONSTANTES GLOBALES PARA POSICION COLUMNAS HOJA INVENTARIO
' ============================================================================================
' Constantes para FILAS de hoja INVENTARIO
Public Const CONST_INVENTARIO_FILA_HEADERS As Integer = 2
' Constantes para COLUMNAS de hoja INVENTARIO
Public Const CONST_INVENTARIO_COLUMNA_NOMBRE As Integer = 2
Public Const CONST_INVENTARIO_COLUMNA_LINK As Integer = 3
Public Const CONST_INVENTARIO_COLUMNA_VISIBLE As Integer = 4
Public Const CONST_INVENTARIO_COLUMNA_FICHERO As Integer = 5
' Constantes para valor HEADERS de hoja INVENTARIO
Public Const CONST_INVENTARIO_HEADER_NOMBRE As String = "Nombre de la Hoja"
Public Const CONST_INVENTARIO_HEADER_LINK As String = "Link a la Hoja"
Public Const CONST_INVENTARIO_HEADER_VISIBLE As String = "Visible/Oculta"
Public Const CONST_INVENTARIO_HEADER_FICHERO As String = "Fichero Fuente"
' Constantes para el valor almacenado en la columna VISIBLE: "OCULTA" o ">>> visible <<<"
Public Const CONST_INVENTARIO_TAG_VISIBLE As String = ">> visible <<"
Public Const CONST_INVENTARIO_TAG_OCULTA As String = "OCULTA"

' ============================================================================================
' CONSTANTES GLOBALES PARA POSICION COLUMNAS HOJA LOG
' ============================================================================================
' Constantes para FILAS de hoja LOG
Public Const CONST_LOG_FILA_HEADERS As Integer = 1
' Constantes para COLUMNAS de hoja LOG
Public Const CONST_LOG_COLUMNA_FECHA_HORA As Integer = 1
Public Const CONST_LOG_COLUMNA_USUARIO As Integer = 2
Public Const CONST_LOG_COLUMNA_TIPO As Integer = 3
Public Const CONST_LOG_COLUMNA_FICHERO As Integer = 4
Public Const CONST_LOG_COLUMNA_HOJA As Integer = 5
Public Const CONST_LOG_COLUMNA_MENSAJE As Integer = 6
' Constantes para valor HEADERS de hoja LOG
Public Const CONST_LOG_HEADER_FECHA_HORA As String = "Fecha/Hora"
Public Const CONST_LOG_HEADER_USUARIO As String = "Usuario"
Public Const CONST_LOG_HEADER_TIPO As String = "Tipo"
Public Const CONST_LOG_HEADER_FICHERO As String = "Fichero"
Public Const CONST_LOG_HEADER_HOJA As String = "Hoja"
Public Const CONST_LOG_HEADER_MENSAJE As String = "Mensaje"

' Variable para hoja de envío anterior
Public gstrPreviaHojaImportacion_Envio As String
' Variable para nombre de copia de hoja de envío anterior
Public gstrPrevDelHojaImportacion_Envio As String


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

'borrame: Public Const CONST_OCULTAR_REPOSITORIO_DELIMITADORES As Boolean = True 'Poner como True si se desea ocultar la hoja
'borrame: Public Const CONST_OCULTAR_HOJA_USERNAME As Boolean = True 'Poner como True si se desea ocultar la hoja

' Variables para celdas que contienen valores originales
Public vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal As String
Public vCelda_Valor_Excel_DecimalSeparator_ValorOriginal As String
Public vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal As String

' Variables para almacenar valores originales leídos
Public vExcel_UseSystemSeparators_ValorOriginal As String
Public vExcel_DecimalSeparator_ValorOriginal As String
Public vExcel_ThousandsSeparator_ValorOriginal As String

' Variables globales para delimitadores
Public vDelimitadorDecimal_HFM As String
Public vDelimitadorMiles_HFM As String


' =============================================================================
' VARIABLES GLOBALES PARA DETECCIÓN DE RANGOS POR PALABRAS CLAVE
' =============================================================================

Public vPalabraClave_PrimeraFila As String
Public vPalabraClave_PrimeraColumna As String
Public vPalabraClave_UltimaFila As String
Public vPalabraClave_UltimaColumna As String


'------------------------------------------------------------------------------
' Procedimiento: VARIABLES Y CONSTANTES PARA LA CONEXION DE SMART VIEW
'------------------------------------------------------------------------------

' Constantes para la Conexion
Public Const CONST_PROVIDER As String = "Hyperion Financial Management"
Public Const CONST_PROVIDER_URL As String = "http://sv3572.logista.local:19000/hfmadf/officeprovider"
Public Const CONST_SERVER_NAME As String = "HFM"
Public Const CONST_APPLICATION_NAME As String = "BUCONS1012"
Public Const CONST_DATABASE_NAME As String = "BUCONS1012"
Public Const CONST_CONNECTION_FRIENDLY_NAME As String = "Conexion_BUCONS1012_Presupuesto"
Public Const CONST_DESCRIPTION As String = "Conexion_BUCONS1012_Presupuesto"

' Constantes para los SmartView > Options > Data Options
Public Const CONST_INDENT_SETTING As Integer = 5
    Public Const CONST_INDENT_NONE As Integer = 0
    Public Const CONST_INDENT_CHILD As Integer = 1
    Public Const CONST_INDENT_PARENT As Integer = 2
Public Const CONST_SUPPRESS_MISSING_SETTING As Integer = 6
Public Const CONST_SUPPRESS_ZERO_SETTING As Integer = 7
Public Const CONST_ENABLE_NOACCESS_MEMBERS_SETTING As Integer = 9
Public Const CONST_ENABLE_REPEATED_MEMBERS_SETTING As Integer = 10
Public Const CONST_ENABLE_INVALID_MEMBERS_SETTING As Integer = 11
Public Const CONST_CELL_DISPLAY_SETTING As Integer = 15
    Public Const CONST_CELL_DISPLAY_SHOW_DATA As Integer = 0
    Public Const CONST_CELL_DISPLAY_SHOW_CALC_STATUS As Integer = 1
    Public Const CONST_CELL_DISPLAY_SHOW_PROCESS_MANAGEMENT As Integer = 2
Public Const CONST_DISPLAY_MEMBER_NAME_SETTING As Integer = 16
    Public Const CONST_DISPLAY_NAME_ONLY As Integer = 0
    Public Const CONST_DISPLAY_AND_DESCRIPTION As Integer = 1
    Public Const CONST_DISPLAY_DESCRIPTION_ONLY As Integer = 2
    
' Variables para las credenciales de SmartView
Public vUsername As String
Public vPassword As String

' Constantes para mostrar o no mensajes durante la ejecución
Public Const CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS As Boolean = False
Public Const CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION As Boolean = False
Public Const CONST_MOSTRAR_MENSAJE_FINAL_SMARTVIEW_CREAR_CONEXION As Boolean = True

Public Const CONST_MOSTRAR_MENSAJES_SMARTVIEW_FIJAR_CONEXION_ACTIVA As Boolean = False
Public Const CONST_MOSTRAR_MENSAJE_FINAL_SMARTVIEW_FIJAR_CONEXION_ACTIVA As Boolean = True

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
    
    'Inicializar hoja de envio anterior
    gstrPreviaHojaImportacion_Envio = ""
    'Inicializar nombre de copia de hoja anterior
    gstrPrevDelHojaImportacion_Envio = ""
    
        
    'Inicializar variables de líneas
    glngLineaInicial_HojaImportacion = 0
    glngLineaFinal_HojaImportacion = 0
    
    'Inicializar nombres de hojas
    gstrNuevaHojaImportacion = ""
    gstrNuevaHojaImportacion_Working = ""
    gstrNuevaHojaImportacion_Envio = ""
    gstrNuevaHojaImportacion_Comprobacion = ""
    
    'Adicional
    'Configuración de importación
    vColumnaInicial_Importacion = "B"        ' Columna B (2)
    vFilaInicial_Importacion = 2             ' Fila 2
    vDelimitador_Importacion = ";"           ' Delimitador por defecto
    
    'Inicializar variables de rango
    vLineaInicial_HojaImportacion = 0
    vLineaFinal_HojaImportacion = 0
    
    'Inicializar palabras clave para detección de rangos
    vPalabraClave_PrimeraFila = "BUDGET_OS"      ' Palabra clave para primera fila
    vPalabraClave_PrimeraColumna = "BUDGET_OS"   ' Palabra clave para primera columna
    vPalabraClave_UltimaFila = "BUDGET_OS"       ' Palabra clave para última fila
    vPalabraClave_UltimaColumna = "M12"          ' Palabra clave para última columna
    
    'Configurar nombre de backup automático
    If gstrPreviaHojaImportacion_Envio <> "" Then
        gstrPrevDelHojaImportacion_Envio = CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO & gstrPreviaHojaImportacion_Envio 'Normalmente CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO="BK_"
    End If
    
    
End Sub

