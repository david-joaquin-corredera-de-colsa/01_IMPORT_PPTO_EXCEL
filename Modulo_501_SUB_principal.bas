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
    
    ThisWorkbook.Save
    
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
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Importaci�n:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion & vbCrLf & vbCrLf & _
           "Esta hoja contendr� los datos importados.", _
           vbInformation, _
           "Hoja de Importaci�n - " & strFuncion
    
    ' Mostrar nombre de la hoja de trabajo
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Trabajo (Working):" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Working & vbCrLf & vbCrLf & _
           "Esta hoja se utilizar� para procesamiento temporal.", _
           vbInformation, _
           "Hoja de Trabajo - " & strFuncion
    
    ' Mostrar nombre de la hoja de env�o
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Env�o:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Envio & vbCrLf & vbCrLf & _
           "Esta hoja contendr� los datos listos para env�o.", _
           vbInformation, _
           "Hoja de Env�o - " & strFuncion
           
    ' Mostrar nombre de la hoja de comprobaci�n
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Comprobaci�n:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Comprobacion & vbCrLf & vbCrLf & _
           "Esta hoja se utilizar� para verificaci�n y control de calidad.", _
           vbInformation, _
           "Hoja de Comprobaci�n - " & strFuncion
           
    '--------------------------------------------------------------------------
    ' 3a. Localizar hoja de env�o anterior
    '--------------------------------------------------------------------------
    lngLineaError = 89
    fun801_LogMessage "Iniciando localizaci�n de hoja de env�o anterior..."
    
    blnResult = F009_Localizar_Hoja_Envio_Anterior()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1011, strFuncion, _
            "La localizaci�n de hoja de env�o anterior no se complet� correctamente"
    End If
    
    Dim vPrefijo_Del_Prev_Envio, vSufijo_Del_Prev_Envio As String
    vPrefijo_Del_Prev_Envio = "Del_Prev_Envio_"
    vSufijo_Del_Prev_Envio = Right(gstrPreviaHojaImportacion_Envio, 15)
    
    gstrPrevDelHojaImportacion_Envio = vPrefijo_Del_Prev_Envio & vSufijo_Del_Prev_Envio
    
    '--------------------------------------------------------------------------
    ' 3b. Copiar hoja de env�o anterior
    '--------------------------------------------------------------------------
    lngLineaError = 90
    fun801_LogMessage "Iniciando copia de hoja de env�o anterior..."
    
    blnResult = F010_Copiar_Hoja_Envio_Anterior()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1012, strFuncion, _
            "La copia de hoja de env�o anterior no se complet� correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Ejecutar proceso de importaci�n (F002)
    '--------------------------------------------------------------------------
    lngLineaError = 91
    fun801_LogMessage "Iniciando proceso de importaci�n..."
    
    blnResult = F002_Importar_Fichero(gstrNuevaHojaImportacion, _
                                     gstrNuevaHojaImportacion_Working, _
                                     gstrNuevaHojaImportacion_Envio, _
                                     gstrNuevaHojaImportacion_Comprobacion)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, _
            "El proceso de importaci�n no se complet� correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 5. Procesar hoja de env�o
    '--------------------------------------------------------------------------
    lngLineaError = 95
    fun801_LogMessage "Iniciando procesamiento de hoja de env�o..."
    
    blnResult = F003_Procesar_Hoja_Envio(gstrNuevaHojaImportacion_Working, _
                                        gstrNuevaHojaImportacion_Envio)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, _
            "El procesamiento de la hoja de env�o no se complet� correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 6. Procesar hoja de comprobaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 97
    fun801_LogMessage "Iniciando procesamiento de hoja de comprobaci�n..."
    
    blnResult = F005_Procesar_Hoja_Comprobacion()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1004, strFuncion, _
            "El procesamiento de la hoja de comprobaci�n NO se complet� correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 6.8.1. SmartView: Crear Conexiones, Actualizar Opciones, Crear AdHoc
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    lngLineaError = 106
    fun801_LogMessage "Creando conexiones para la hoja " & gstrPrevDelHojaImportacion_Envio & "..."
    
    Call M002_SmartView_Paso_01_CrearConexiones_EstablecerOpciones_CrearAdHoc(gstrPrevDelHojaImportacion_Envio)
    
    '--------------------------------------------------------------------------
    ' 6.8.2. SmartView: localizo final inicial y final
    '                   (como en F007_Copiar_Datos_de_Comprobacion_a_Envio)
    '                   y voy editando cada celda de datos (con valor en blanco)
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    lngLineaError = 107
    fun801_LogMessage "Copiando datos de hoja de comprobaci�n a hoja de env�o..."
    
    blnResult = F007_Preparar_Datos_para_Borrado(gstrPrevDelHojaImportacion_Envio)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, _
            "La copia de datos de comprobaci�n a env�o NO se complet� correctamente"
    End If
    
    
    '--------------------------------------------------------------------------
    ' 7. SmartView: Crear Conexiones, Actualizar Opciones, Crear AdHoc
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    Call M002_SmartView_Paso_01_CrearConexiones_EstablecerOpciones_CrearAdHoc(gstrNuevaHojaImportacion_Envio)
    
    '--------------------------------------------------------------------------
    ' 8.1. Copiar datos de comprobaci�n a env�o
    '--------------------------------------------------------------------------
    lngLineaError = 107
    fun801_LogMessage "Copiando datos de hoja de comprobaci�n a hoja de env�o..."
    
    blnResult = F007_Copiar_Datos_de_Comprobacion_a_Envio(gstrNuevaHojaImportacion_Envio, _
                                                          gstrNuevaHojaImportacion_Comprobacion)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, _
            "La copia de datos de comprobaci�n a env�o NO se complet� correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 8.2. SmartView: Enviar Datos / Submit
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    Call M003_SmartView_Paso_02_Submit(gstrNuevaHojaImportacion_Envio)
    'Call M003_SmartView_Paso_02_Submit_without_Refresh(gstrNuevaHojaImportacion_Envio)
    
    '--------------------------------------------------------------------------
    ' 9. Restaurar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 103
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Restaurar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1007, strFuncion, _
            "Error en la detecci�n de delimitadores"
    End If
    
    ThisWorkbook.Save
    
    '--------------------------------------------------------------------------
    ' 9a. NUEVA FUNCIONALIDAD: Limpieza de Hojas Historicas
    '--------------------------------------------------------------------------
    lngLineaError = 109
    fun801_LogMessage "Iniciando proceso de limpieza de hojas historicas..."
    
    blnResult = Function_Return_Integer_to_Boolean(F011_Limpieza_Hojas_Historicas())
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1010, strFuncion, _
            "El proceso de limpieza de hojas historicas NO se complet� correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 9b. NUEVA FUNCIONALIDAD: Limpieza de Hojas Historicas
    '--------------------------------------------------------------------------
    lngLineaError = 110
    fun801_LogMessage "Iniciando proceso de inventariar hojas..."
    
    blnResult = Function_Return_Integer_to_Boolean(F012_Inventariar_Hojas())
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1010, strFuncion, _
            "El proceso de inventariar NO se complet� correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 10. Proceso completado
    '--------------------------------------------------------------------------
    lngLineaError = 111
    fun801_LogMessage "Proceso principal completado correctamente"
    
    MsgBox "El proceso se ha completado correctamente." & vbCrLf & vbCrLf & _
           "- Hojas creadas: 4" & vbCrLf & _
           "- Datos importados en: " & gstrNuevaHojaImportacion & vbCrLf & _
           "- Rango de datos: " & glngLineaInicial_HojaImportacion & " a " & _
           glngLineaFinal_HojaImportacion & vbCrLf & _
           "- Hoja de comprobaci�n preparada: " & gstrNuevaHojaImportacion_Comprobacion, _
           vbInformation, _
           "�xito - " & strFuncion
               
    ThisWorkbook.Worksheets("00_Ejecutar_Procesos").Select
    
    ThisWorkbook.Save
           
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


