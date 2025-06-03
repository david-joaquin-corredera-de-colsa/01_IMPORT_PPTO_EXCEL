Attribute VB_Name = "Modulo_501_SUB_principal"
Option Explicit

'******************************************************************************
' Módulo: M001_Ejecutar_Proceso_Principal
' Fecha y Hora de Creación: 2025-05-26 05:39:34 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripción:
' Este módulo contiene el procedimiento principal que coordina la ejecución
' de los procesos de importación y gestión de datos.
'******************************************************************************

Public Sub M001_Ejecutar_Proceso_Principal()

    '--------------------------------------------------------------------------
    ' Variables para control de errores y seguimiento
    '--------------------------------------------------------------------------
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    
    ' Inicialización
    strFuncion = "M001_Ejecutar_Proceso_Principal"
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 0. Configuración inicial del entorno
    '--------------------------------------------------------------------------
    lngLineaError = 44
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Inicializar variables globales
    Call InitializeGlobalVariables
    
    fun801_LogMessage "Iniciando proceso principal..."
    
    '--------------------------------------------------------------------------
    ' 1. Ejecución de comprobaciones iniciales (F000)
    '--------------------------------------------------------------------------
    lngLineaError = 54
    fun801_LogMessage "Ejecutando comprobaciones iniciales..."
    
    blnResult = F000_Comprobaciones_Iniciales()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "Las comprobaciones iniciales no se completaron correctamente"
    End If
    
    
    '--------------------------------------------------------------------------
    ' 2. Creacion de hojas de importacion (F001)
    '--------------------------------------------------------------------------
    lngLineaError = 55
    fun801_LogMessage "Creando hojas de importacion..."

    blnResult = F001_Crear_hojas_de_Importacion()

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "Las hojas de importacion no se crearon correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 2a. Detectar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 61
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Detectar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    ThisWorkbook.Save
    
    '--------------------------------------------------------------------------
    ' 2b. Forzar delimitadores Especificos en el sistema  |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 62
    Call fun801_LogMessage("Forzando delimitadores Especificos en el sistema", False)

    blnResult = F004_Forzar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Mostrar información de las hojas creadas
    '--------------------------------------------------------------------------
    lngLineaError = 66
    
    ' Mostrar nombre de la hoja de importación
    MsgBox "Hoja de Importación:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion & vbCrLf & vbCrLf & _
           "Esta hoja contendrá los datos importados.", _
           vbInformation, _
           "Hoja de Importación - " & strFuncion
    
    ' Mostrar nombre de la hoja de trabajo
    MsgBox "Hoja de Trabajo (Working):" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Working & vbCrLf & vbCrLf & _
           "Esta hoja se utilizará para procesamiento temporal.", _
           vbInformation, _
           "Hoja de Trabajo - " & strFuncion
    
    ' Mostrar nombre de la hoja de envío
    MsgBox "Hoja de Envío:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Envio & vbCrLf & vbCrLf & _
           "Esta hoja contendrá los datos listos para envío.", _
           vbInformation, _
           "Hoja de Envío - " & strFuncion
           
    ' Mostrar nombre de la hoja de comprobación
    MsgBox "Hoja de Comprobación:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Comprobacion & vbCrLf & vbCrLf & _
           "Esta hoja se utilizará para verificación y control de calidad.", _
           vbInformation, _
           "Hoja de Comprobación - " & strFuncion
           
    '--------------------------------------------------------------------------
    ' 4. Ejecutar proceso de importación (F002)
    '--------------------------------------------------------------------------
    lngLineaError = 91
    fun801_LogMessage "Iniciando proceso de importación..."
    
    blnResult = F002_Importar_Fichero(gstrNuevaHojaImportacion, _
                                     gstrNuevaHojaImportacion_Working, _
                                     gstrNuevaHojaImportacion_Envio, _
                                     gstrNuevaHojaImportacion_Comprobacion)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, _
            "El proceso de importación no se completó correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 5. NUEVA FUNCIONALIDAD: Procesar hoja de envío
    '--------------------------------------------------------------------------
    lngLineaError = 95
    fun801_LogMessage "Iniciando procesamiento de hoja de envío..."
    
    blnResult = F003_Procesar_Hoja_Envio(gstrNuevaHojaImportacion_Working, _
                                        gstrNuevaHojaImportacion_Envio)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, _
            "El procesamiento de la hoja de envío no se completó correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 6. NUEVA FUNCIONALIDAD: Procesar hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 97
    fun801_LogMessage "Iniciando procesamiento de hoja de comprobación..."
    
    blnResult = F005_Procesar_Hoja_Comprobacion()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1004, strFuncion, _
            "El procesamiento de la hoja de comprobación NO se completó correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 7. NUEVA FUNCIONALIDAD: Procesar hoja de comprobación
    '--------------------------------------------------------------------------
'    lngLineaError = 99
'    fun801_LogMessage "Gestion de las conexiones de SmartView (actualizacion de la conexion y conexion activa)..."
'
'    blnResult = SmartView_CreateConnection
'
'    If Not blnResult Then
'        Err.Raise ERROR_BASE_IMPORT + 1005, strFuncion, _
'            "La conexion NO se actualizo correctamente y " & vbCrLf & "NO se establecio como conexion activa."
'    End If
    
    '--------------------------------------------------------------------------
    ' 8.1. NUEVA FUNCIONALIDAD: Actualizar Data Options / Member Options en SmartView
    '--------------------------------------------------------------------------
'    lngLineaError = 101
'    fun801_LogMessage "DataOptions/MemberOptions en SmartView"
'
'    blnResult = SmartView_Options_DataOptions_Estandar(gstrNuevaHojaImportacion_Envio)
'
'    If Not blnResult Then
'        Err.Raise ERROR_BASE_IMPORT + 1006, strFuncion, _
'            "Las DataOptions/Member Options NO se actualizaron correctamente en SmartView" & vbCrLf & "--"
'    End If


    '--------------------------------------------------------------------------
    ' 8.2. NUEVA FUNCIONALIDAD: Actualizar Data Options / Member Options en SmartView
    '--------------------------------------------------------------------------
'    lngLineaError = 105
'    fun801_LogMessage "Refresh en SmartView"
'
'    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
'    blnResult = SmartView_Retrieve(gstrNuevaHojaImportacion_Envio)
'
'    If Not blnResult Then
'        Err.Raise ERROR_BASE_IMPORT + 1008, strFuncion, _
'            "El Refresh NO se hizo correctamente en SmartView" & vbCrLf & "--"
'    End If
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    'Call M002_SmartView_Paso_01
    Call M002_SmartView_Paso_01_v2(gstrNuevaHojaImportacion_Envio)
    
    '--------------------------------------------------------------------------
    ' 9. Restaurar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 103
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Restaurar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1007, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    ThisWorkbook.Save
    
    '--------------------------------------------------------------------------
    ' 10. Proceso completado
    '--------------------------------------------------------------------------
    lngLineaError = 110
    fun801_LogMessage "Proceso principal completado correctamente"
    
    MsgBox "El proceso se ha completado correctamente." & vbCrLf & vbCrLf & _
           "- Hojas creadas: 4" & vbCrLf & _
           "- Datos importados en: " & gstrNuevaHojaImportacion & vbCrLf & _
           "- Rango de datos: " & glngLineaInicial_HojaImportacion & " a " & _
           glngLineaFinal_HojaImportacion & vbCrLf & _
           "- Hoja de comprobación preparada: " & gstrNuevaHojaImportacion_Comprobacion, _
           vbInformation, _
           "Éxito - " & strFuncion
           
CleanExit:
    '--------------------------------------------------------------------------
    ' 7. Restauración del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

GestorErrores:
    ' Construcción del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Origen: " & Err.Source & vbCrLf & _
                      "Descripción: " & Err.Description
    
    ' Registro del error
    fun801_LogMessage strMensajeError, True
    
    ' Mostrar mensaje al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    
    ' Asegurar que se restaura la configuración de Excel
    Resume CleanExit
End Sub


