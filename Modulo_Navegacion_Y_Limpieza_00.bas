Attribute VB_Name = "Modulo_Navegacion_Y_Limpieza_00"

' =============================================================================
' MODULO: Modulo_Configuracion.bas
' PROYECTO: IMPORTAR_DATOS_PRESUPUESTO
' AUTOR: david-joaquin-corredera-de-colsa
' FECHA CREACION: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Modulo de configuracion global y constantes del proyecto
' COMPATIBILIDAD: Excel 97, Excel 2003, Excel 2007, Excel 365
' REPOSITORIO: OneDrive, SharePoint, Teams compatible
' =============================================================================

Option Explicit

' =============================================================================
' CONSTANTES GLOBALES DEL PROYECTO
' =============================================================================

' Configuracion de hojas del sistema
Public Const HOJA_INICIAL As String = "00_Ejecutar_Procesos"
Public Const HOJA_INVENTARIO As String = "01_Inventario"
Public Const HOJA_LOG As String = "02_Log"
Public Const HOJA_REPORT As String = "09_Report_PL"
Public Const HOJA_USERNAME As String = "05_Username"
Public Const HOJA_DELIMITADORES As String = "06_Delimitadores_Originales"

' Configuracion de limpieza de hojas historicas
Public Const MAX_HOJAS_ENVIO_VISIBLES As Integer = 5
Public Const PREFIJO_IMPORT_WORKING As String = "Import_Working_"
Public Const PREFIJO_IMPORT_COMPROB As String = "Import_Comprob_"
Public Const PREFIJO_IMPORT_ENVIO As String = "Import_Envio_"
Public Const PREFIJO_DEL_PREV_ENVIO As String = "Del_Prev_Envio_"
Public Const PREFIJO_IMPORT_GENERICO As String = "Import_"
Public Const LONGITUD_IMPORT_GENERICO As Integer = 22

' Codigos de error personalizados
Public Const ERROR_LIBRO_NO_DISPONIBLE As Integer = 1001
Public Const ERROR_HOJA_NO_ENCONTRADA As Integer = 1002
Public Const ERROR_ACCESO_HOJA_INVENTARIO As Integer = 3001
Public Const ERROR_NO_ESPECIFICADO_LIMPIEZA As Integer = 2999
Public Const ERROR_NO_ESPECIFICADO_INVENTARIO As Integer = 3999
Public Const ERROR_NO_ESPECIFICADO_GENERAL As Integer = 9999

' =============================================================================
' FUNCION: Obtener_Version_Proyecto
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Devuelve la version actual del proyecto
' RETORNO: String (version)
' =============================================================================
Public Function Obtener_Version_Proyecto() As String
    Obtener_Version_Proyecto = "1.0.0 - 2025-06-03 15:18:26 UTC"
End Function

' =============================================================================
' FUNCION: Validar_Configuracion_Sistema
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Valida que todas las hojas requeridas existan
' RETORNO: Boolean (True=todas existen, False=faltan hojas)
' =============================================================================
Public Function Validar_Configuracion_Sistema() As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim vHojasRequeridas(1 To 6) As String
    Dim i As Integer
    Dim vHojaEncontrada As Boolean
    Dim j As Integer
    
    ' Lista de hojas requeridas para el funcionamiento
    vHojasRequeridas(1) = HOJA_INICIAL
    vHojasRequeridas(2) = HOJA_INVENTARIO
    vHojasRequeridas(3) = HOJA_LOG
    vHojasRequeridas(4) = HOJA_REPORT
    vHojasRequeridas(5) = HOJA_USERNAME
    vHojasRequeridas(6) = HOJA_DELIMITADORES
    
    ' Verificar cada hoja requerida
    For i = 1 To 6
        vHojaEncontrada = False
        For j = 1 To ThisWorkbook.Worksheets.Count
            If StrComp(ThisWorkbook.Worksheets(j).Name, vHojasRequeridas(i), vbTextCompare) = 0 Then
                vHojaEncontrada = True
                Exit For
            End If
        Next j
        
        If Not vHojaEncontrada Then
            Validar_Configuracion_Sistema = False
            Exit Function
        End If
    Next i
    
    Validar_Configuracion_Sistema = True
    Exit Function
    
ErrorHandler:
    Validar_Configuracion_Sistema = False
    
End Function
