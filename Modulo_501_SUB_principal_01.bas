Attribute VB_Name = "Modulo_501_SUB_principal_01"
Option Explicit

Public Sub M001_Ejecutar_Proceso_Principal()

    '******************************************************************************
    ' Módulo: M001_Ejecutar_Proceso_Principal
    ' Fecha y Hora de Creación: 2025-05-26 05:39:34 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Este módulo contiene el procedimiento principal que coordina la ejecución
    ' de los procesos de importación y gestión de datos.
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Variables para control de errores y seguimiento
    '--------------------------------------------------------------------------
    '*********************************
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    '*********************************
    Dim blnInventarioActualizado As Boolean
    '*********************************
    Dim vCredencialesObtenidas As Boolean
    Dim vConexionCreada As Boolean
    Dim vConexionBorrada As Boolean
    Dim vConexionActivada As Boolean
    
    Dim vReturn_SmartView_Retrieve As Boolean
    Dim vReturn_SmartView_Establecer_Options_Estandar As Integer
    Dim vReturn_SmartView_Submit As Integer
    Dim vReturn_SmartView_Submit_without_Refresh As Integer

    
    'Variable para habilitar/deshabilitar partes de esta SUB
    Dim vEnabled_Parts As Boolean
    'vEnabled_Parts = True
    'If vEnabled_Parts Then
    'End If 'vEnabled_Parts Then
    
    ' Inicialización
    strFuncion = "M001_Ejecutar_Proceso_Principal" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "M001_Ejecutar_Proceso_Principal"
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
            "Las comprobaciones iniciales " & vbCrLf & " no se completaron correctamente"
    End If
    
    
    '--------------------------------------------------------------------------
    ' 2. Creacion de hojas de importacion (F001)
    '--------------------------------------------------------------------------
    lngLineaError = 55
    fun801_LogMessage "Creando hojas de importacion..."

    blnResult = F001_Crear_hojas_de_Importacion()

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, _
            "Las hojas de importacion " & vbCrLf & " no se crearon correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 2a. Detectar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 61
    Call fun801_LogMessage("Detectando delimitadores del sistema", False)

    blnResult = F004_Detectar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, _
            "Error en la detección de delimitadores"
    End If
    
    'ThisWorkbook.Save
    
    '--------------------------------------------------------------------------
    ' 2b. Forzar delimitadores Especificos en el sistema  |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 62
    Call fun801_LogMessage("Forzando delimitadores Especificos en el sistema", False)

    blnResult = F004_Forzar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1004, strFuncion, _
            "Error al forzar delimitadores especificos en el sistema"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Mostrar información de las hojas creadas
    '--------------------------------------------------------------------------
    lngLineaError = 66
    
    ' Mostrar nombre de la hoja de importación
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Importación:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion & vbCrLf & vbCrLf & _
           "Esta hoja contendrá los datos importados.", _
           vbInformation, _
           "Hoja de Importación - " & strFuncion
    
    ' Mostrar nombre de la hoja de trabajo
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Trabajo (Working):" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Working & vbCrLf & vbCrLf & _
           "Esta hoja se utilizará para procesamiento temporal.", _
           vbInformation, _
           "Hoja de Trabajo - " & strFuncion
    
    ' Mostrar nombre de la hoja de envío
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Envío:" & vbCrLf & vbCrLf & _
           gstrNuevaHojaImportacion_Envio & vbCrLf & vbCrLf & _
           "Esta hoja contendrá los datos listos para envío.", _
           vbInformation, _
           "Hoja de Envío - " & strFuncion
           
    ' Mostrar nombre de la hoja de comprobación
    If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then MsgBox "Hoja de Comprobación:" & vbCrLf & vbCrLf & _
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
        Err.Raise ERROR_BASE_IMPORT + 1005, strFuncion, _
            "El proceso de importación no se completó correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 5. Procesar hoja de envío
    '--------------------------------------------------------------------------
    lngLineaError = 95
    fun801_LogMessage "Iniciando procesamiento de hoja de envío..."
    
    Dim vScenario_HEnvio As String
    Dim vYear_HEnvio As String
    Dim vEntity_HEnvio As String
    
    blnResult = F003_Procesar_Hoja_Envio(gstrNuevaHojaImportacion_Working, _
                                        gstrNuevaHojaImportacion_Envio, vScenario_HEnvio, vYear_HEnvio, vEntity_HEnvio)
    
    'MsgBox "vScenario_HEnvio=" & vScenario_HEnvio
    'MsgBox "vYear_HEnvio=" & vYear_HEnvio
    'MsgBox "vEntity_HEnvio=" & vEntity_HEnvio
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1006, strFuncion, _
            "El procesamiento de la hoja de envío no se completó correctamente"
    End If
        
    '--------------------------------------------------------------------------
    ' 6. Procesar hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 97
    fun801_LogMessage "Iniciando procesamiento de hoja de comprobación..."
    
    blnResult = F005_Procesar_Hoja_Comprobacion()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1007, strFuncion, _
            "El procesamiento de la hoja de comprobación NO se completó correctamente"
    End If

    '--------------------------------------------------------------------------
    ' 6a. Localizar hoja de envío anterior
    '--------------------------------------------------------------------------
    lngLineaError = 89
    fun801_LogMessage "Iniciando localización de hoja de envío anterior..."

    blnResult = F009_Localizar_Hoja_Envio_Anterior(vScenario_HEnvio, vYear_HEnvio, vEntity_HEnvio)

    'Variables para borrar los contenidos de la columna en la que etiquetamos si se envio el dato correcto o no (vColumnaEtiqueta)
    Dim vColumnaEtiqueta As Integer
    vColumnaEtiqueta = 1
    Dim vHojaEnvioAntigua As Worksheet
    Set vHojaEnvioAntigua = ThisWorkbook.Worksheets(gstrPreviaHojaImportacion_Envio)
    'Borramos los contenidos de la columna en la que etiquetamos si se envio el dato correcto o no (vColumnaEtiqueta)
    vHojaEnvioAntigua.Columns(vColumnaEtiqueta).ClearContents
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1008, strFuncion, _
            "La localización de hoja de envío anterior no se completó correctamente"
    End If

    Dim vSufijo_Del_Prev_Envio As String
    vSufijo_Del_Prev_Envio = Right(gstrPreviaHojaImportacion_Envio, 15)

    gstrPrevDelHojaImportacion_Envio = CONST_PREFIJO_HOJA_X_BORRAR_ENVIO_PREVIO & vSufijo_Del_Prev_Envio 'Normalmente CONST_PREFIJO_HOJA_X_BORRAR_ENVIO_PREVIO = "Del_Prev_Envio_"
    
    '--------------------------------------------------------------------------
    ' 6b. Copiar hoja de envío anterior
    '--------------------------------------------------------------------------
    lngLineaError = 997005
    fun801_LogMessage "Iniciando copia de hoja de envío anterior..."

    blnResult = F010_Copiar_Hoja_Envio_Anterior()

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, _
            "La copia de hoja de envío anterior no se completó correctamente"
    End If

    'Variables para borrar los contenidos de la columna en la que etiquetamos si se envio el dato correcto o no (vColumnaEtiqueta)
    Dim vHojaBorradoDatosAntiguos As Worksheet
    Set vHojaBorradoDatosAntiguos = ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio)
    'Borramos los contenidos de la columna en la que etiquetamos si se envio el dato correcto o no (vColumnaEtiqueta)
    vHojaBorradoDatosAntiguos.Columns(vColumnaEtiqueta).ClearContents

    '--------------------------------------------------------------------------
    ' 6.8.0.    SmartView:  Para todas las hojas
    '                       gstrPrevDelHojaImportacion_Envio y gstrNuevaHojaImportacion_Envio
    '                       Pedimos credenciales
    '--------------------------------------------------------------------------
    
    lngLineaError = 997006
    fun801_LogMessage "Solicitando credenciales para usar en todas las hojas" & "..."
    
    'Pedimos las credenciales
    vCredencialesObtenidas = Pedir_Credenciales(vUsername, vPassword)
    
    If Not vCredencialesObtenidas Then
        Err.Raise ERROR_BASE_IMPORT + 1010, strFuncion, _
            "Las credenciales no se obtuvieron correctamente"
    End If
    
        
    '--------------------------------------------------------------------------
    ' 6.8.1.    SmartView:  PARA TODAS LAS HOJAS
    '           Creamos ciertas variables necesarias para la creación de CONEXIONES, etc.
    '                       y les damos valor (tomamos ese valor de las constantes globales o de las credenciales)
    '--------------------------------------------------------------------------

    lngLineaError = 997007
    fun801_LogMessage "Modificando el valor de ciertas variables necesarias" & "..."

    'MsgBox "Username=" & vUsername & "|Password=" & vPassword
    Dim vConnection_Username As String: vConnection_Username = vUsername
    Dim vConnection_Password As String: vConnection_Password = vPassword
    Dim vConnection_Provider As String: vConnection_Provider = CONST_PROVIDER
    Dim vConnection_URL As String: vConnection_URL = CONST_PROVIDER_URL
    Dim vConnection_Server As String: vConnection_Server = CONST_SERVER_NAME
    Dim vConnection_Application As String: vConnection_Application = CONST_APPLICATION_NAME
    Dim vConnection_Database As String: vConnection_Database = CONST_DATABASE_NAME
    Dim vConnection_Name As String: vConnection_Name = CONST_CONNECTION_FRIENDLY_NAME
    Dim vConnection_Description As String: vConnection_Description = CONST_DESCRIPTION
    Dim vConnection_Create_MostrarMensajes As Boolean: vConnection_Create_MostrarMensajes = CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION
    Dim vConnection_Create_MostrarMensajeFinal As Boolean: vConnection_Create_MostrarMensajeFinal = CONST_MOSTRAR_MENSAJE_FINAL_SMARTVIEW_CREAR_CONEXION
    
    '--------------------------------------------------------------------------
    ' 6.8.2.    SmartView:  PARA TODAS LAS HOJAS
    '                       Creamos la conexion
    '--------------------------------------------------------------------------

    lngLineaError = 997008
    fun801_LogMessage "Creando la conexion " & Chr(34) & vConnection_Name & Chr(34) & " ..."
    
    vConexionCreada = SmartView_Create_Connection(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal)
    
    If Not vConexionCreada Then
        Err.Raise ERROR_BASE_IMPORT + 1011, strFuncion, _
            "La conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & " NO se creó correctamente"
    End If
    
    '--------------------------------------------------------------------------
    ' 6.8.3.    SmartView:  Para la hoja del ultimo envío gstrPrevDelHojaImportacion_Envio
    '                       Fijamos la conexion como activa
    '--------------------------------------------------------------------------

    lngLineaError = 997009
    fun801_LogMessage "Fijando la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "como conexión activa ..." & _
        "para la hoja " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34)
    
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    
    Dim vConnection_FijarActiva_MostrarMensajes As Boolean: vConnection_Create_MostrarMensajes = CONST_MOSTRAR_MENSAJES_SMARTVIEW_FIJAR_CONEXION_ACTIVA
    Dim vConnection_FijarActiva_MostrarMensajeFinal As Boolean: vConnection_Create_MostrarMensajeFinal = CONST_MOSTRAR_MENSAJE_FINAL_SMARTVIEW_FIJAR_CONEXION_ACTIVA
    
    'La establecemos como activa para la hoja gstrPrevDelHojaImportacion_Envio
    vConexionActivada = SmartView_SetActiveConnection_x_Sheet(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, gstrPrevDelHojaImportacion_Envio)

    If Not vConexionActivada Then
        Err.Raise ERROR_BASE_IMPORT + 1012, strFuncion, _
            "La conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "no fijó como activa para la hoja " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34)
    End If

    '--------------------------------------------------------------------------
    ' 6.8.4.    SmartView:  Para la hoja del ultimo envío gstrPrevDelHojaImportacion_Envio
    '                       Fijamos opciones de Datos y Formato
    '--------------------------------------------------------------------------
    
    lngLineaError = 997010
    fun801_LogMessage "Fijando opciones de Datos y Formato para la la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "sobre la hoja " & _
         Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34) & " ..."

    '01. Establecemos las opciones estandar sobre la hoja gstrPrevDelHojaImportacion_Envio
    vReturn_SmartView_Establecer_Options_Estandar = SmartView_Establecer_Options_Estandar(gstrPrevDelHojaImportacion_Envio)
    
    '02. Hacemos Retrieve/Refresh en la hoja gstrPrevDelHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(gstrPrevDelHojaImportacion_Envio)
    

    '--------------------------------------------------------------------------
    ' 6.8.5. SmartView: para la hoja del último envío gstrPrevDelHojaImportacion_Envio
    '                   localizo fila inicial, fila final,
    '                   columna inicial, columna final,
    '                   y voy editando cada celda de datos (con valor en blanco)
    '                   para que luego nos deje hacer su envío correctamente
    '                   (=nos deje borrar el dato que enviamos la vez anterior)
    '--------------------------------------------------------------------------

    lngLineaError = 997011
    fun801_LogMessage "Copiando datos de hoja previa de importación " & Chr(34) & gstrPreviaHojaImportacion_Envio & Chr(34) & _
        "a hoja de envío para borrado de datos antigüos " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34) & vbCrLf & "y dejando sus celdas como 0/blank"

    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70

    blnResult = F007_Preparar_Datos_para_Borrado(gstrPreviaHojaImportacion_Envio, gstrPrevDelHojaImportacion_Envio) 'Crear esta nueva funcion 20250604

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1013, strFuncion, _
            "La copia de datos y subsiguiente 'borrado' de la hoja " & Chr(34) & gstrPreviaHojaImportacion_Envio & Chr(34) & vbCrLf & _
            "a la hoja " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34) & vbCrLf & "NO se completó correctamente"
    End If

    '--------------------------------------------------------------------------
    ' 6.8.6. SmartView: para la hoja del último envío gstrPrevDelHojaImportacion_Envio
    '                   Enviar Datos / Submit
    '--------------------------------------------------------------------------

    lngLineaError = 997012
    fun801_LogMessage "Enviando a HFM datos (vacíos/blank) con la hoja previa de importación " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34) & vbCrLf & "a HFM ..."
    
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70

    Dim vMensajeBorradoConExito As String
    vMensajeBorradoConExito = "Se borraron los datos antiguos con exito."
    
    vReturn_SmartView_Submit = SmartView_Submit(gstrPrevDelHojaImportacion_Envio, vMensajeBorradoConExito)
    'vReturn_SmartView_Submit_without_Refresh = SmartView_Submit_without_Refresh(gstrPrevDelHojaImportacion_Envio,vMensajeBorradoConExito)
    
    '--------------------------------------------------------------------------
    ' 6.8.7. SmartView: para la hoja del último envío gstrPrevDelHojaImportacion_Envio
    '                   Hacer Retrieve/Refresh de Datos
    '--------------------------------------------------------------------------

    lngLineaError = 997013
    fun801_LogMessage "Haciendo retrieve/refresh de datos de HFM" & vbCrLf & "con la hoja previa de importación " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34) & " ..."
    
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrPrevDelHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    'Hacemos Retrieve/Refresh en la hoja gstrPrevDelHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(gstrPrevDelHojaImportacion_Envio)

    If Not vReturn_SmartView_Retrieve Then
        Err.Raise ERROR_BASE_IMPORT + 1014, strFuncion, _
            "Para la hoja " & Chr(34) & gstrPrevDelHojaImportacion_Envio & Chr(34) & vbCrLf & "no consiguio hacer retrieve ..."
    End If


    '--------------------------------------------------------------------------
    ' 7.1. SmartView:   Para la nueva hoja de envío gstrNuevaHojaImportacion_Envio
    '                   Fijamos la conexion como activa
    '--------------------------------------------------------------------------

    lngLineaError = 997014
    fun801_LogMessage "Fijando la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "como conexión activa ..." & _
        "para la hoja " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34)
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70

    vConexionActivada = SmartView_SetActiveConnection_x_Sheet(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, gstrNuevaHojaImportacion_Envio)

    '--------------------------------------------------------------------------
    ' 7.2. SmartView:   Para la nueva hoja de envío gstrNuevaHojaImportacion_Envio
    '                   Fijamos opciones de Datos y Formato
    '--------------------------------------------------------------------------
    
    lngLineaError = 997015
    fun801_LogMessage "Fijando opciones de Datos y Formato para la la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "sobre la hoja " & _
         Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34) & " ..."

    '01. Establecemos las opciones estandar sobre la hoja gstrNuevaHojaImportacion_Envio
    vReturn_SmartView_Establecer_Options_Estandar = SmartView_Establecer_Options_Estandar(gstrNuevaHojaImportacion_Envio)
    
    '02. Hacemos Retrieve/Refresh en la hoja gstrNuevaHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(gstrNuevaHojaImportacion_Envio)

    '--------------------------------------------------------------------------
    ' 7.3. SmartView: para la hoja de envío gstrNuevaHojaImportacion_Envio
    '                   localizo fila inicial, fila final,
    '                   columna inicial, columna final,
    '                   y voy editando cada celda de datos (mantengo su valor)
    '                   para que luego nos deje hacer su envío correctamente
    '--------------------------------------------------------------------------
        
    lngLineaError = 997016
    fun801_LogMessage "Copiando datos de hoja de comprobación " & Chr(34) & gstrNuevaHojaImportacion_Comprobacion & Chr(34) & vbCrLf & _
        " a hoja de envío " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34)

    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70

    Dim vRangoCalculo As String '20250615
    blnResult = F007_Copiar_Datos_de_Comprobacion_a_Envio(gstrNuevaHojaImportacion_Comprobacion, _
                                                          gstrNuevaHojaImportacion_Envio, vRangoCalculo) '20250615

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1014, strFuncion, _
            "La copia de datos de la hoja de comprobación " & Chr(34) & gstrNuevaHojaImportacion_Comprobacion & Chr(34) & vbCrLf & _
            "a la hoja de envío " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34) & vbCrLf & _
            "NO se completó correctamente"
    End If

    '--------------------------------------------------------------------------
    ' 7.4. SmartView:   para la hoja de envío gstrNuevaHojaImportacion_Envio
    '                   Enviar Datos / Submit
    '--------------------------------------------------------------------------

    lngLineaError = 997017
    fun801_LogMessage "Enviando a HFM los datos de la hoja de envío " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34) & vbCrLf & "a HFM ..."

    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70

    Dim vMensajeCargaConExito As String
    vMensajeCargaConExito = "Se caragron los datos nuevos con exito."

    vReturn_SmartView_Submit = SmartView_Submit(gstrNuevaHojaImportacion_Envio, vMensajeCargaConExito)
    'vReturn_SmartView_Submit_without_Refresh = SmartView_Submit_without_Refresh(gstrNuevaHojaImportacion_Envio,vMensajeCargaConExito)


    '--------------------------------------------------------------------------
    ' 7.5. SmartView:   para la hoja de envío gstrNuevaHojaImportacion_Envio
    '                   Hacer Retrieve/Refresh de Datos
    '--------------------------------------------------------------------------

    lngLineaError = 997018
    fun801_LogMessage "Haciendo retrieve/refresh de los datos de HFM" & vbCrLf & "con la nueva hoja de importación " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34) & " ..."
    
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70

    
    'Primero hacer refresh de gstrNuevaHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(gstrNuevaHojaImportacion_Envio)
    
    'Segundo Calculamos
    Dim vResultadoHypCalculate, vResultadoHypForceCalculate As Long '20250615
    'vResultadoHypCalculate = HypCalculate(gstrNuevaHojaImportacion_Envio, vRangoCalculo) '20250615
    vResultadoHypCalculate = HypForceCalculate(gstrNuevaHojaImportacion_Envio, vRangoCalculo) '20250615
    'MsgBox "vResultadoHypCalculate=" & vResultadoHypCalculate '20250615
    'MsgBox "vResultadoHypForceCalculate=" & vResultadoHypForceCalculate '20250615

    'Tercero, volvemos a hacer refresh de gstrNuevaHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(gstrNuevaHojaImportacion_Envio)

    If Not vReturn_SmartView_Retrieve Then
        Err.Raise ERROR_BASE_IMPORT + 1015, strFuncion, _
            "Para la hoja " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34) & vbCrLf & "no consiguio hacer retrieve ..."
    End If
    
    
    '--------------------------------------------------------------------------
    ' 7.6. SmartView: para la hoja de envío gstrNuevaHojaImportacion_Envio
    '                   localizo fila inicial, fila final,
    '                   columna inicial, columna final,
    '                   y comparo cada celda de datos
    '                   con su valor en la hoja gstrNuevaHojaImportacion_Comprobacion
    '                   y si es OK etiqueto la linea en VERDE
    '                   pero si es NO-OK etiqueto la linea en ROJO
    '--------------------------------------------------------------------------
        
    lngLineaError = 997019
    fun801_LogMessage "Comprobando los datos que pretendiamos cargar en HFM > " & Chr(34) & gstrNuevaHojaImportacion_Comprobacion & Chr(34) & vbCrLf & _
        "con los datos realmente cargados en HFMm > " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34)

    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Select
    ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio).Activate
    ActiveWindow.Zoom = 70
    
    Dim vScenario_xPL As String
    Dim vYear_xPL As String
    Dim vEntity_xPL As String
    
    blnResult = F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio(gstrNuevaHojaImportacion_Comprobacion, _
                                                          gstrNuevaHojaImportacion_Envio, vScenario_xPL, vYear_xPL, vEntity_xPL)

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1016, strFuncion, _
            "La copia de datos de la hoja de comprobación " & Chr(34) & gstrNuevaHojaImportacion_Comprobacion & Chr(34) & vbCrLf & _
            "a la hoja de envío " & Chr(34) & gstrNuevaHojaImportacion_Envio & Chr(34) & vbCrLf & _
            "NO se completó correctamente"
    End If

    'MsgBox "vScenario_xPL = " & vScenario_xPL
    'MsgBox "vYear_xPL = " & vYear_xPL
    'MsgBox "vEntiy_xPL = " & vEntity_xPL

    '--------------------------------------------------------------------------
    ' 7.6.0.    La hoja de la PL AdHoc (CONST_HOJA_REPORT_PL_AH / vReport_PL_AH_Name)
    '           si esta oculta la mostramos
    '--------------------------------------------------------------------------
    
    Dim vReport_PL_AH_Name As String
    vReport_PL_AH_Name = CONST_HOJA_REPORT_PL_AH
    
    lngLineaError = 997020
    fun801_LogMessage "Mostrando / haciendo visible la hoja oculta..." & _
        Chr(34) & vReport_PL_AH_Name & Chr(34)

    blnResult = fun823_MostrarHojaSiOculta(vReport_PL_AH_Name)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1017, strFuncion, _
            "La hoja oculta " & Chr(34) & vReport_PL_AH_Name & Chr(34) & vbCrLf & "no se pudo mostrar / hacer visible " & Chr(34) & vReport_PL_AH_Name & Chr(34)
    End If
    
    '--------------------------------------------------------------------------
    ' 7.6.1.    Seleccionamos la hoja de la PL AdHoc (CONST_HOJA_REPORT_PL_AH / vReport_PL_AH_Name)
    '           y establecemos la conexion como activa para esa hoja
    '--------------------------------------------------------------------------
           
    
    lngLineaError = 997021
    fun801_LogMessage "Fijando la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "como conexión activa ..." & _
        "para la hoja " & Chr(34) & vReport_PL_AH_Name & Chr(34)
    
    
    ThisWorkbook.Worksheets(vReport_PL_AH_Name).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(vReport_PL_AH_Name).Select
    ThisWorkbook.Worksheets(vReport_PL_AH_Name).Activate
    ActiveWindow.Zoom = 70
        
    'La establecemos como activa para la hoja gstrPrevDelHojaImportacion_Envio
    vConexionActivada = SmartView_SetActiveConnection_x_Sheet(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal, vReport_PL_AH_Name)

    If Not vConexionActivada Then
        Err.Raise ERROR_BASE_IMPORT + 1018, strFuncion, _
            "La conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "no fijó como activa para la hoja " & Chr(34) & vReport_PL_AH_Name & Chr(34)
    End If
    
    '--------------------------------------------------------------------------
    ' 7.6.2.    SmartView:  Para la hoja de la PL AH (CONST_HOJA_REPORT_PL_AH / vReport_PL_AH_Name)
    '                       Fijamos opciones de Datos y Formato
    '--------------------------------------------------------------------------
    
    lngLineaError = 997022
    fun801_LogMessage "Fijando opciones de Datos y Formato para la la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "sobre la hoja " & _
         Chr(34) & vReport_PL_AH_Name & Chr(34) & " ..."

    '01. Establecemos las opciones estandar sobre la hoja gstrPrevDelHojaImportacion_Envio
    vReturn_SmartView_Establecer_Options_Estandar = SmartView_Establecer_Options_Estandar(vReport_PL_AH_Name)
        
    If Not vReturn_SmartView_Establecer_Options_Estandar Then
        Err.Raise ERROR_BASE_IMPORT + 1019, strFuncion, _
            "Para la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "no se pudieron establecer opciones estandar de formato " & _
            vbCrLf & " sobre la hoja " & Chr(34) & vReport_PL_AH_Name & Chr(34)
    End If
    
    '--------------------------------------------------------------------------
    ' 7.6.3.    La PL AdHoc siempre va a tener la misma estructura
    '           por eso podemos trabajar con una estructura fija de filas/columnas almacenada en constantes/variables
    '           sin tener que detectar donde se encuentra realmente cada dimension (ya está pre-localizada cada dimensión)
    '
    '           Sabiendo esto, vamos a actualizar Scenario, Year, Entity para la hoja PL AdHoc
    '--------------------------------------------------------------------------
    
    lngLineaError = 997023
    fun801_LogMessage "Modificando el Report de PL - versión AdHoc " & vbCrLf & _
        "sobre la hoja " & Chr(34) & vReport_PL_AH_Name & Chr(34) & " ..." & vbCrLf & _
        "modificando Scenario, Year, Entity"
        
    Dim vFilaScenario, vFilaYear, vFilaEntity As Integer
    vFilaScenario = CONST_FILA_SCENARIO: vFilaYear = CONST_FILA_YEAR: vFilaEntity = CONST_FILA_ENTITY
    
    Dim vColumnaInicialHeaders, vColumnaFinalHeaders As Integer
    vColumnaInicialHeaders = CONST_COLUMNA_INICIAL_HEADERS: vColumnaFinalHeaders = CONST_COLUMNA_FINAL_HEADERS
    '20250609: seguir aqui
    
    blnResult = Modificar_Scenario_Year_Entity_en_hoja_PLAH(vReport_PL_AH_Name, _
                                                            vFilaScenario, vFilaYear, vFilaEntity, _
                                                            vColumnaInicialHeaders, vColumnaFinalHeaders, _
                                                            vScenario_xPL, vYear_xPL, vEntity_xPL)

    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1020, strFuncion, _
            "Para la hoja " & Chr(34) & vReport_PL_AH_Name & Chr(34) & vbCrLf & "hubo un error al modificar Scenario, Year, Entity ..."
    End If
    
    '--------------------------------------------------------------------------
    ' 7.6.4.  Hacemos el Retrieve para que nos reconozca el AdHoc
    '--------------------------------------------------------------------------
    'MsgBox "Pre 7.6.4."
    lngLineaError = 997024
    fun801_LogMessage "Haciendo Retrieve/Refresh para la la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "sobre la hoja " & _
         Chr(34) & vReport_PL_AH_Name & Chr(34) & " ..."

    '02. Hacemos Retrieve/Refresh en la hoja gstrPrevDelHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(vReport_PL_AH_Name)

    If Not vReturn_SmartView_Retrieve Then
        Err.Raise ERROR_BASE_IMPORT + 1021, strFuncion, _
            "Para la hoja " & Chr(34) & vReport_PL_AH_Name & Chr(34) & vbCrLf & "no consiguio hacer retrieve tras modificar Scenario, Year, Entity ..."
    End If
    
    '--------------------------------------------------------------------------
    ' 7.6.5.  Habilitar seleccion de Actividad, Negocio, Entity '20250609: seguir aqui
    '--------------------------------------------------------------------------

    lngLineaError = 997025
    fun801_LogMessage "Hciendo Retrieve/Refresh para la la conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & "sobre la hoja " & _
         Chr(34) & vReport_PL_AH_Name & Chr(34) & " ..."

    '01. Habilitar el Member Selection > seguir aqui '20250609: seguir aqui
    
    '02. Hacemos Retrieve/Refresh en la hoja gstrPrevDelHojaImportacion_Envio
    vReturn_SmartView_Retrieve = SmartView_Retrieve(vReport_PL_AH_Name)
    
    'MsgBox "Post 7.6.5."


    '--------------------------------------------------------------------------
    ' 8. Restaurar delimitadores Originales del sistema |||||||
    '--------------------------------------------------------------------------
    lngLineaError = 997027
    Call fun801_LogMessage("Restaurando delimitadores originales del sistema", False)

    blnResult = F004_Restaurar_Delimitadores_en_Excel()
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1023, strFuncion, _
            "Error en la restauracion de los delimitadores originales del sistema"
    End If
    
    GoTo CleanExit ' Asegura que se ejecuta la limpieza aunque no haya error

CleanExit:

    '--------------------------------------------------------------------------
    ' 9.0.    SmartView:  borramos la conexion
    '--------------------------------------------------------------------------
    
    lngLineaError = 997026
    fun801_LogMessage "Borrando la conexion " & Chr(34) & vConnection_Name & Chr(34) & " ..."
    
    vConexionBorrada = SmartView_Delete_Connection(vConnection_Username, vConnection_Password, vConnection_Provider, vConnection_URL, vConnection_Server, _
        vConnection_Application, vConnection_Database, vConnection_Name, vConnection_Description, _
        vConnection_Create_MostrarMensajes, vConnection_Create_MostrarMensajeFinal)
    
    If Not vConexionBorrada Then
        Err.Raise ERROR_BASE_IMPORT + 1022, strFuncion, _
            "La conexion " & Chr(34) & vConnection_Name & Chr(34) & vbCrLf & " NO se BORRÓ correctamente"
    End If

    '--------------------------------------------------------------------------
    ' 9.a. Ordenamos las hojas que quedan en el libro
    '--------------------------------------------------------------------------
    lngLineaError = 997028
    fun801_LogMessage "Ordenando las hojas del libro ..."
    
    blnResult = Ordenar_Hojas()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'ordenar' las hojas"
    End If


    '--------------------------------------------------------------------------
    ' 9.b. Comprobamos si el inventario esta actualizado
    '--------------------------------------------------------------------------
    lngLineaError = 997029
    fun801_LogMessage "Comprobando si el inventario esta actualizado ..."
    
    blnInventarioActualizado = Inventario_Actualizado_Si_No()
    
    If blnInventarioActualizado Then
        
        fun801_LogMessage "El inventario de hojas estaba actualizado (no será necesario actualizarlo ahora mismo)"
        
    Else
        
        fun801_LogMessage "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        'MsgBox "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        '--------------------------------------------------------------------------
        ' 9.c. Limpiamos las hojas "tecnicas" historicas
        '   borramos las que no son "Import_Envio_"
        '   dejamos visibles solo las ultimas X (5) hojas de "Import_Envio_"
        '   más recientes
        '--------------------------------------------------------------------------
        lngLineaError = 997030
        fun801_LogMessage "Limpiando las hojas tecnicas históricas ..."
        
        blnResult = F011_Limpieza_Hojas_Historicas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente la limpeza de hojas tecnicas históricas"
        End If
            
        '--------------------------------------------------------------------------
        ' 9.d. Actualizamos la hoja 01_Inventario
        '   Listando todas las hojas existentes en el libro
        '--------------------------------------------------------------------------
        lngLineaError = 997031
        fun801_LogMessage "Inventariando las hojas del libro ..."
        
        blnResult = F012_Inventariar_Hojas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente el inventario de hojas"
        End If
        'MsgBox "Acabamos de actualizar el inventario de hojas"
    End If
    
    '--------------------------------------------------------------------------
    ' 9.e. Regresamos a la hoja Principal del libro
    '--------------------------------------------------------------------------

    lngLineaError = 997032
    fun801_LogMessage "Ejecutando apertura de libro > abriendo la hoja inicial ..."
    
    blnResult = F010_Abrir_Hoja_Inicial()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se ejecutó correctamente la apertura de libro > apertura de hoja inicial"
    End If
    
    '--------------------------------------------------------------------------
    ' 10. Proceso completado
    '--------------------------------------------------------------------------
    lngLineaError = 997033
    fun801_LogMessage "Proceso principal completado correctamente"

    MsgBox "El proceso se ha completado correctamente." & vbCrLf & vbCrLf & _
           "Éxito - " & strFuncion

    
    '--------------------------------------------------------------------------
    ' 7. Restauración del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    '--------------------------------------------------------------------------
    ' 8. Exit Sub
    '--------------------------------------------------------------------------
    
    ThisWorkbook.Save
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


