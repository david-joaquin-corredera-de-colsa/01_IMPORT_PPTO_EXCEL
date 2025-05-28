Attribute VB_Name = "Modulo_501_SUB_principal"
Option Explicit

'******************************************************************************
' M�dulo: M001_Ejecutar_Proceso_Principal
' Fecha y Hora de Creaci�n: 2025-05-26 05:39:34 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripci�n:
' Este m�dulo contiene el procedimiento principal que coordina la ejecuci�n
' de los procesos de importaci�n y gesti�n de datos.
'******************************************************************************

Public Sub M001_Ejecutar_Proceso_Principal()

    '--------------------------------------------------------------------------
    ' Variables para control de errores y seguimiento
    '--------------------------------------------------------------------------
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    
    ' Inicializaci�n
    strFuncion = "M001_Ejecutar_Proceso_Principal"
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 0. Configuraci�n inicial del entorno
    '--------------------------------------------------------------------------
    lngLineaError = 44
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Inicializar variables globales
    Call InitializeGlobalVariables
    
    fun801_LogMessage "Iniciando proceso principal..."
    
    '--------------------------------------------------------------------------
    ' 1. Ejecuci�n de comprobaciones iniciales (F000)
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
            "Error en la detecci�n de delimitadores"
    End If
    
    '--------------------------------------------------------------------------
    ' 2b. Forzar delimitadores Especificos en el sistema  |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 62
    Call fun801_LogMessage("Forzando delimitadores Especificos en el sistema", False)

    blnResult = F004_Forzar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error en la detecci�n de delimitadores"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Mostrar informaci�n de las hojas creadas
    '--------------------------------------------------------------------------
    lngLineaError = 66
    
    ' Mostrar nombre de la hoja de importaci�n
    MsgBox "Hoja de Importaci�n:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion & vbCrLf & vbCrLf & _
           "Esta hoja contendr� los datos importados.", _
           vbInformation, _
           "Hoja de Importaci�n - " & strFuncion
    
    ' Mostrar nombre de la hoja de trabajo
    MsgBox "Hoja de Trabajo (Working):" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Working & vbCrLf & vbCrLf & _
           "Esta hoja se utilizar� para procesamiento temporal.", _
           vbInformation, _
           "Hoja de Trabajo - " & strFuncion
    
    ' Mostrar nombre de la hoja de env�o
    MsgBox "Hoja de Env�o:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Envio & vbCrLf & vbCrLf & _
           "Esta hoja contendr� los datos listos para env�o.", _
           vbInformation, _
           "Hoja de Env�o - " & strFuncion
           
    '--------------------------------------------------------------------------
    ' 4. Ejecutar proceso de importaci�n (F002)
    '--------------------------------------------------------------------------
    lngLineaError = 91
    fun801_LogMessage "Iniciando proceso de importaci�n..."
    
    blnResult = F002_Importar_Fichero(gstrNuevaHojaImportacion, _
                                     gstrNuevaHojaImportacion_Working, _
                                     gstrNuevaHojaImportacion_Envio)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, _
            "El proceso de importaci�n no se complet� correctamente"
    End If
    
    
'    '--------------------------------------------------------------------------
'    ' 5. Detectar delimitadores del sistema
'    '--------------------------------------------------------------------------
'    lngLineaError = 92
'    Call fun801_LogMessage("Detectando delimitadores del sistema", False)
'
'    blnResult = F003_Delimitador_Excel()
'    If Not blnResult Then
'        Err.Raise ERROR_BASE_PROCESS + 6, strSubName, _
'            "Error en la detecci�n de delimitadores"
'    End If
    
    '--------------------------------------------------------------------------
    ' 9. Restaurar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 95
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Restaurar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error en la detecci�n de delimitadores"
    End If
    
    
    '--------------------------------------------------------------------------
    ' 10. Proceso completado
    '--------------------------------------------------------------------------
    lngLineaError = 110
    fun801_LogMessage "Proceso principal completado correctamente"
    
    MsgBox "El proceso se ha completado correctamente." & vbCrLf & vbCrLf & _
           "- Hojas creadas: 3" & vbCrLf & _
           "- Datos importados en: " & gstrNuevaHojaImportacion & vbCrLf & _
           "- Rango de datos: " & glngLineaInicial_HojaImportacion & " a " & _
           glngLineaFinal_HojaImportacion, _
           vbInformation, _
           "�xito - " & strFuncion


CleanExit:
    '--------------------------------------------------------------------------
    ' 7. Restauraci�n del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Origen: " & Err.Source & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    ' Registro del error
    fun801_LogMessage strMensajeError, True
    
    ' Mostrar mensaje al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    
    ' Asegurar que se restaura la configuraci�n de Excel
    Resume CleanExit
End Sub


