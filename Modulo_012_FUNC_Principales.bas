Attribute VB_Name = "Modulo_012_FUNC_Principales"
Option Explicit
Public Function F000_Comprobaciones_Iniciales() As Boolean
    

    '******************************************************************************
    ' M�dulo: F000_Comprobaciones_Iniciales
    ' Fecha y Hora de Creaci�n: 2025-05-26 09:32:08 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Esta funci�n realiza las comprobaciones iniciales necesarias y crea las hojas
    ' requeridas para el proceso de importaci�n.
    '
    ' Pasos:
    ' 1. Inicializaci�n de variables globales
    ' 2. Validaci�n y creaci�n de hojas base (Procesos, Inventario, Log)
    ' 3. Generaci�n de nombres para nuevas hojas de importaci�n
    ' 4. Creaci�n de hojas de importaci�n
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para fechas y nombres de hojas
    Dim strFechaHoraIsoActual As String
    Dim strFechaHoraIsoNuevaHojaImportacion As String
    Dim strPrefijoHojaImportacion As String
    Dim strPrefijoHojaImportacion_Working As String
    Dim strPrefijoHojaImportacion_Envio As String
    
    ' Inicializaci�n
    strFuncion = "F000_Comprobaciones_Iniciales"
    F000_Comprobaciones_Iniciales = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializar variables globales
    '--------------------------------------------------------------------------
    lngLineaError = 50
    Call InitializeGlobalVariables
    
    '--------------------------------------------------------------------------
    ' 2. Validar/Crear hojas base
    '--------------------------------------------------------------------------
    lngLineaError = 57
    ' Validar/Crear hoja Ejecutar Procesos
    If Not fun802_SheetExists(gstrHoja_EjecutarProcesos) Then
        If Not F002_Crear_Hoja(gstrHoja_EjecutarProcesos) Then
            Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, _
                "Error al crear la hoja " & gstrHoja_EjecutarProcesos
        End If
    End If
    
    ' Validar/Crear hoja Inventario
    If Not fun802_SheetExists(gstrHoja_Inventario) Then
        If Not F002_Crear_Hoja(gstrHoja_Inventario) Then
            Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, _
                "Error al crear la hoja " & gstrHoja_Inventario
        End If
    End If
    
    ' Validar/Crear hoja Log
    If Not fun802_SheetExists(gstrHoja_Log) Then
        If Not F002_Crear_Hoja(gstrHoja_Log) Then
            Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
                "Error al crear la hoja " & gstrHoja_Log
        End If
    End If
    
    
    
    ' Proceso completado exitosamente
    F000_Comprobaciones_Iniciales = True
    fun801_LogMessage "Comprobaciones iniciales completadas con �xito"
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F000_Comprobaciones_Iniciales = False
End Function


Public Function F001_Crear_hojas_de_Importacion() As Boolean
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para fechas y nombres de hojas
    Dim strFechaHoraIsoActual As String
    Dim strFechaHoraIsoNuevaHojaImportacion As String
    Dim strPrefijoHojaImportacion As String
    Dim strPrefijoHojaImportacion_Working As String
    Dim strPrefijoHojaImportacion_Envio As String
    
    ' Inicializaci�n
    strFuncion = "F001_Crear_hojas_de_Importacion"
    F001_Crear_hojas_de_Importacion = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializar variables globales
    '--------------------------------------------------------------------------
    lngLineaError = 51
    Call InitializeGlobalVariables
   
    '--------------------------------------------------------------------------
    ' 3. Generar nombres para nuevas hojas
    '--------------------------------------------------------------------------
    lngLineaError = 85
    ' Generar timestamp ISO
    strFechaHoraIsoActual = Format(Now(), "yyyymmdd_hhmmss")
    strFechaHoraIsoNuevaHojaImportacion = strFechaHoraIsoActual
    
    ' Definir prefijos
    strPrefijoHojaImportacion = "Import_"
    strPrefijoHojaImportacion_Working = "Import_Working_"
    strPrefijoHojaImportacion_Envio = "Import_Envio_"
    
    ' Generar nombres completos (variables globales)
    gstrNuevaHojaImportacion = strPrefijoHojaImportacion & strFechaHoraIsoNuevaHojaImportacion
    gstrNuevaHojaImportacion_Working = strPrefijoHojaImportacion_Working & strFechaHoraIsoNuevaHojaImportacion
    gstrNuevaHojaImportacion_Envio = strPrefijoHojaImportacion_Envio & strFechaHoraIsoNuevaHojaImportacion
    
    '--------------------------------------------------------------------------
    ' 4. Crear hojas de importaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 102
    ' Crear hoja de importaci�n
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion) Then
        Err.Raise ERROR_BASE_IMPORT + 4, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion
    End If
    
    ' Crear hoja de trabajo
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion_Working) Then
        Err.Raise ERROR_BASE_IMPORT + 5, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion_Working
    End If
    
    ' Crear hoja de env�o
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion_Envio
    End If
    
    ' Proceso completado exitosamente
    F001_Crear_hojas_de_Importacion = True
    fun801_LogMessage "Creacion de hojas de importacion completada con �xito"
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F001_Crear_hojas_de_Importacion = False
End Function



Public Function F002_Importar_Fichero(ByVal vNuevaHojaImportacion As String, _
                                    ByVal vNuevaHojaImportacion_Working As String, _
                                    ByVal vNuevaHojaImportacion_Envio As String) As Boolean
    
    '******************************************************************************
    ' M�dulo: F002_Importar_Fichero
    ' Fecha y Hora de Creaci�n: 2025-05-26 10:50:40 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Funci�n para importar ficheros de texto a Excel, manteniendo el formato original
    ' en la hoja de importaci�n y procesando los datos en la hoja de trabajo.
    '
    ' Pasos:
    ' 1. Limpieza de hojas destino (Importaci�n, Working, Env�o)
    ' 2. Selecci�n de archivo mediante cuadro de di�logo
    ' 3. Importaci�n de datos sin procesar a hoja de importaci�n
    ' 4. Copia de datos a hoja de trabajo
    ' 5. Procesamiento en hoja de trabajo:
    '    - Detecci�n de rango de datos
    '    - Conversi�n de texto a columnas con formatos espec�ficos
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas y rangos
    Dim wsImport As Worksheet
    Dim wsWorking As Worksheet
    Dim wsEnvio As Worksheet
    Dim rngConversion As Range
    
    ' Variables para importaci�n
    Dim strFilePath As String
    Dim lngCol As Long
    
    ' Inicializaci�n
    strFuncion = "F002_Importar_Fichero"
    F002_Importar_Fichero = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Limpiar hojas destino
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando proceso de importaci�n", False, "", ""
    
    ' Limpiar hoja de importaci�n
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion
    End If
    
    ' Limpiar hoja working
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion_Working
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion_Working) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion_Working
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion_Working
    End If
    
    ' Limpiar hoja env�o
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion_Envio
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion_Envio) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion_Envio
        Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion_Envio
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Seleccionar archivo
    '--------------------------------------------------------------------------
    lngLineaError = 71
    fun801_LogMessage "Solicitando selecci�n de archivo al usuario", False, "", ""
    strFilePath = fun802_SeleccionarArchivo("�Qu� fichero desea importar?")
    
    If strFilePath = "" Then
        fun801_LogMessage "No se seleccion� ning�n archivo", True, "", ""
        Err.Raise ERROR_BASE_IMPORT + 4, strFuncion, _
            "No se seleccion� ning�n archivo"
    End If
    
    fun801_LogMessage "Archivo seleccionado para importar", False, strFilePath, vNuevaHojaImportacion
    
    '--------------------------------------------------------------------------
    ' 3. Importar datos sin procesar
    '--------------------------------------------------------------------------
    lngLineaError = 81
    fun801_LogMessage "Iniciando importaci�n de archivo", False, strFilePath, vNuevaHojaImportacion
    Set wsImport = ThisWorkbook.Worksheets(vNuevaHojaImportacion)
    
    If Not fun803_ImportarArchivo(wsImport, strFilePath, _
                               vColumnaInicial_Importacion, _
                               vFilaInicial_Importacion) Then
        fun801_LogMessage "Error en la importaci�n", True, strFilePath, vNuevaHojaImportacion
        Err.Raise ERROR_BASE_IMPORT + 5, strFuncion, _
            "Error al importar el archivo"
    End If
    
    fun801_LogMessage "Archivo importado correctamente", False, strFilePath, vNuevaHojaImportacion
    
    '--------------------------------------------------------------------------
    ' 4. Copiar datos a hoja working
    '--------------------------------------------------------------------------
    lngLineaError = 95
    fun801_LogMessage "Copiando datos a hoja de trabajo", False, strFilePath, vNuevaHojaImportacion_Working
    Set wsWorking = ThisWorkbook.Worksheets(vNuevaHojaImportacion_Working)
    
    ' Copiar datos
    wsImport.UsedRange.Copy wsWorking.Range(vColumnaInicial_Importacion & vFilaInicial_Importacion)
    fun801_LogMessage "Datos copiados correctamente", False, strFilePath, vNuevaHojaImportacion_Working
    
    '--------------------------------------------------------------------------
    ' 5. Procesar datos en hoja working
    '--------------------------------------------------------------------------
    lngLineaError = 104
    ' Detectar rango de datos
    fun801_LogMessage "Detectando rango de datos", False, strFilePath, vNuevaHojaImportacion_Working
    If Not fun804_DetectarRangoDatos(wsWorking, _
                                  vLineaInicial_HojaImportacion, _
                                  vLineaFinal_HojaImportacion) Then
        fun801_LogMessage "Error al detectar rango de datos", True, strFilePath, vNuevaHojaImportacion_Working
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error al detectar el rango de datos"
    End If
    
    fun801_LogMessage "Rango detectado: " & vLineaInicial_HojaImportacion & " a " & vLineaFinal_HojaImportacion, _
                      False, strFilePath, vNuevaHojaImportacion_Working
    
    ' Seleccionar rango para conversi�n
    Set rngConversion = wsWorking.Range( _
        vColumnaInicial_Importacion & vLineaInicial_HojaImportacion & ":" & _
        vColumnaInicial_Importacion & vLineaFinal_HojaImportacion)
    
    ' Convertir texto a columnas con formatos espec�ficos
    lngLineaError = 120
    fun801_LogMessage "Iniciando conversi�n texto a columnas", False, strFilePath, vNuevaHojaImportacion_Working
    
    With rngConversion
        .TextToColumns _
            Destination:=.Cells(1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=False, _
            Semicolon:=(vDelimitador_Importacion = ";"), _
            Comma:=(vDelimitador_Importacion = ","), _
            Space:=(vDelimitador_Importacion = " "), _
            Other:=True, _
            OtherChar:=IIf(vDelimitador_Importacion <> ";" And _
                          vDelimitador_Importacion <> "," And _
                          vDelimitador_Importacion <> " ", _
                          vDelimitador_Importacion, "")
        
        ' Configurar formato de columnas
        lngCol = Range(vColumnaInicial_Importacion & "1").Column
        
        ' Columnas 1-11 como texto
        wsWorking.Range(wsWorking.Cells(vLineaInicial_HojaImportacion, lngCol), _
                       wsWorking.Cells(vLineaFinal_HojaImportacion, lngCol + 10)).NumberFormat = "@"
        
        ' Columnas 12-23 como General
        wsWorking.Range(wsWorking.Cells(vLineaInicial_HojaImportacion, lngCol + 11), _
                       wsWorking.Cells(vLineaFinal_HojaImportacion, lngCol + 22)).NumberFormat = "General"
    End With
    
    fun801_LogMessage "Conversi�n texto a columnas completada", False, strFilePath, vNuevaHojaImportacion_Working
    
    ' Proceso completado exitosamente
    fun801_LogMessage "Proceso de importaci�n completado con �xito", False, strFilePath, vNuevaHojaImportacion_Working
    F002_Importar_Fichero = True
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, strFilePath, IIf(Len(vNuevaHojaImportacion_Working) > 0, _
                                                              vNuevaHojaImportacion_Working, _
                                                              vNuevaHojaImportacion)
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F002_Importar_Fichero = False
End Function




Public Function F004_Detectar_Delimitadores_en_Excel() As Boolean
    
    ' =============================================================================
    ' FUNCI�N PRINCIPAL: F004_Detectar_Delimitadores_en_Excel
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-05-26 17:43:59 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    ' Descripci�n: Detecta y almacena los delimitadores de Excel actuales
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables globales con valores por defecto
    ' 2. Verificar si existe la hoja de delimitadores originales
    ' 3. Si no existe, crear la hoja y dejarla visible
    ' 4. Si existe, verificar su visibilidad y hacerla visible si est� oculta
    ' 5. Limpiar el contenido de la hoja una vez visible
    ' 6. Configurar headers en las celdas especificadas (B2, B3, B4)
    ' 7. Detectar configuraci�n actual de delimitadores de Excel:
    '    - Use System Separators (True/False)
    '    - Decimal Separator (car�cter)
    '    - Thousands Separator (car�cter)
    ' 8. Almacenar valores detectados en variables globales
    ' 9. Escribir valores en la hoja de delimitadores (C2, C3, C4)
    ' 10. Verificar variable global vOcultarRepostiorioDelimitadores
    ' 11. Si es True, ocultar la hoja creada/actualizada
    ' 12. Manejo exhaustivo de errores con informaci�n detallada
    '
    ' Par�metros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    ' Control de errores con n�mero de l�nea
    On Error GoTo ErrorHandler
    
    ' Variables locales
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim hojaExiste As Boolean
    Dim i As Integer
    Dim lineaError As Long
    
    ' Inicializar resultado como exitoso
    F004_Detectar_Delimitadores_en_Excel = True
    
    ' ==========================================================================
    ' PASO 1: INICIALIZAR VARIABLES GLOBALES CON VALORES POR DEFECTO
    ' ==========================================================================
    lineaError = 100
    
    ' Nombre de la hoja donde se almacenar�n los delimitadores originales
    vHojaDelimitadoresExcelOriginales = "06_Delimitadores_Originales"
    
    ' Celdas para los headers (t�tulos)
    vCelda_Header_Excel_UseSystemSeparators = "B2"
    vCelda_Header_Excel_DecimalSeparator = "B3"
    vCelda_Header_Excel_ThousandsSeparator = "B4"
    
    ' Celdas para los valores detectados
    vCelda_Valor_Excel_UseSystemSeparators = "C2"
    vCelda_Valor_Excel_DecimalSeparator = "C3"
    vCelda_Valor_Excel_ThousandsSeparator = "C4"
    
    ' Variables para almacenar los valores detectados (inicialmente vac�as)
    vExcel_UseSystemSeparators = ""
    vExcel_DecimalSeparator = ""
    vExcel_ThousandsSeparator = ""
    
    lineaError = 110
    
    ' ==========================================================================
    ' PASO 2: OBTENER REFERENCIA AL LIBRO ACTUAL
    ' ==========================================================================
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If wb Is Nothing Then
        F004_Detectar_Delimitadores_en_Excel = False
        Exit Function
    End If
    
    lineaError = 120
    
    ' ==========================================================================
    ' PASO 3: VERIFICAR SI EXISTE LA HOJA DE DELIMITADORES ORIGINALES
    ' ==========================================================================
    
    hojaExiste = fun801_VerificarExistenciaHoja(wb, vHojaDelimitadoresExcelOriginales)
    
    lineaError = 130
    
    ' ==========================================================================
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEG�N CORRESPONDA
    ' ==========================================================================
    
    If Not hojaExiste Then
        ' La hoja no existe, crearla y dejarla visible
        Set ws = fun802_CrearHojaDelimitadores(wb, vHojaDelimitadoresExcelOriginales)
        If ws Is Nothing Then
            F004_Detectar_Delimitadores_en_Excel = False
            Exit Function
        End If
        ' La hoja reci�n creada ya est� visible por defecto
    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
        Set ws = wb.Worksheets(vHojaDelimitadoresExcelOriginales)
        
        ' Verificar si est� oculta y hacerla visible si es necesario
        Call fun803_HacerHojaVisible(ws)
    End If
    
    lineaError = 140
    
    ' ==========================================================================
    ' PASO 5: LIMPIAR CONTENIDO DE LA HOJA (AHORA QUE EST� VISIBLE)
    ' ==========================================================================
    
    Call fun804_LimpiarContenidoHoja(ws)
    
    lineaError = 150
    
    ' ==========================================================================
    ' PASO 6: CONFIGURAR HEADERS EN LAS CELDAS ESPECIFICADAS
    ' ==========================================================================
    
    ' Header para Use System Separators en B2
    ws.Range(vCelda_Header_Excel_UseSystemSeparators).Value = "Excel Use System Separators"
    
    ' Header para Decimal Separator en B3
    ws.Range(vCelda_Header_Excel_DecimalSeparator).Value = "Excel Decimals"
    
    ' Header para Thousands Separator en B4
    ws.Range(vCelda_Header_Excel_ThousandsSeparator).Value = "Excel Thousands"
    
    lineaError = 160
    
    ' ==========================================================================
    ' PASO 7: DETECTAR CONFIGURACI�N ACTUAL DE DELIMITADORES DE EXCEL
    ' ==========================================================================
    
    ' Detectar Use System Separators
    vExcel_UseSystemSeparators = fun805_DetectarUseSystemSeparators()
    
    ' Detectar Decimal Separator
    vExcel_DecimalSeparator = fun806_DetectarDecimalSeparator()
    
    ' Detectar Thousands Separator
    vExcel_ThousandsSeparator = fun807_DetectarThousandsSeparator()
    
    lineaError = 170
    
    ' ==========================================================================
    ' PASO 8: ALMACENAR VALORES DETECTADOS EN LA HOJA
    ' ==========================================================================
    
    ' Almacenar Use System Separators en C2
    ws.Range(vCelda_Valor_Excel_UseSystemSeparators).Value = vExcel_UseSystemSeparators
    
    ' Almacenar Decimal Separator en C3
    ws.Range(vCelda_Valor_Excel_DecimalSeparator).Value = vExcel_DecimalSeparator
    
    ' Almacenar Thousands Separator en C4
    ws.Range(vCelda_Valor_Excel_ThousandsSeparator).Value = vExcel_ThousandsSeparator
    
    lineaError = 180
    
    ' ==========================================================================
    ' PASO 9: VERIFICAR SI DEBE OCULTAR LA HOJA
    ' ==========================================================================
    
    ' Verificar la variable global vOcultarRepostiorioDelimitadores
    If vOcultarRepostiorioDelimitadores = True Then
        ' Ocultar la hoja de delimitadores
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Detectar_Delimitadores_en_Excel - " & Now()
            ' Nota: No es un error cr�tico, el proceso puede continuar
        End If
    End If
    lineaError = 190
    
    ' ==========================================================================
    ' PASO 10: FINALIZACI�N EXITOSA
    ' ==========================================================================
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Detectar_Delimitadores_en_Excel = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: F004_Detectar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now() & vbCrLf & _
                   "USUARIO: david-joaquin-corredera-de-colsa"
    
    ' Mostrar mensaje de error (comentar si no se desea)
    ' MsgBox mensajeError, vbCritical, "Error en Detecci�n de Delimitadores"
    
    ' Log del error para debugging
    Debug.Print mensajeError
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
    
End Function


Public Function F004_Forzar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCI�N: F004_Forzar_Delimitadores_en_Excel
    ' PROP�SITO: Fuerza los delimitadores decimal y de miles en Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' PAR�METROS: Ninguno
    ' RETORNA: Boolean (True = �xito, False = error)
    '
    ' RESUMEN DE PASOS:
    ' 1. Inicializaci�n de variables globales si est�n vac�as
    ' 2. Verificaci�n de compatibilidad del sistema
    ' 3. Backup de configuraci�n actual del usuario
    ' 4. Aplicaci�n de nuevos delimitadores usando Application.International
    ' 5. Verificaci�n de aplicaci�n correcta
    ' 6. Manejo exhaustivo de errores con informaci�n detallada
    ' 7. Retorno de estado de �xito/fallo
    ' =============================================================================

    ' Variables de control de errores
    Dim strFuncionActual As String
    Dim strTipoError As String
    Dim lngLineaError As Long
    
    ' Variables de trabajo
    Dim strDelimitadorDecimalAnterior As String
    Dim strDelimitadorMilesAnterior As String
    Dim blnConfiguracionCambiada As Boolean
    
    ' Inicializaci�n
    strFuncionActual = "F004_Forzar_Delimitadores_en_Excel"
    F004_Forzar_Delimitadores_en_Excel = False
    blnConfiguracionCambiada = False
    
    On Error GoTo ErrorHandler
    
    ' =========================================================================
    ' PASO 1: Inicializaci�n de variables globales
    ' =========================================================================
    lngLineaError = 50
    Call fun801_InicializarVariablesGlobales
    
    ' =========================================================================
    ' PASO 2: Verificaci�n de compatibilidad
    ' =========================================================================
    lngLineaError = 60
    If Not fun802_VerificarCompatibilidad() Then
        strTipoError = "Error de compatibilidad del sistema"
        GoTo ErrorHandler
    End If
    
    ' =========================================================================
    ' PASO 3: Backup de configuraci�n actual
    ' =========================================================================
    lngLineaError = 70
    Call fun803_ObtenerConfiguracionActual(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
    
    ' =========================================================================
    ' PASO 4: Aplicaci�n de nuevos delimitadores
    ' =========================================================================
    lngLineaError = 80
    If fun804_AplicarNuevosDelimitadores() Then
        blnConfiguracionCambiada = True
        
        ' =====================================================================
        ' PASO 5: Verificaci�n de aplicaci�n correcta
        ' =====================================================================
        lngLineaError = 90
        If fun805_VerificarAplicacionDelimitadores() Then
            F004_Forzar_Delimitadores_en_Excel = True
        Else
            strTipoError = "Error en verificaci�n de delimitadores aplicados"
            GoTo ErrorHandler
        End If
    Else
        strTipoError = "Error al aplicar nuevos delimitadores"
        GoTo ErrorHandler
    End If
    
    Exit Function

' =============================================================================
' CONTROL DE ERRORES EXHAUSTIVO
' =============================================================================
ErrorHandler:
    ' Restaurar configuraci�n anterior si se cambi�
    If blnConfiguracionCambiada Then
        On Error Resume Next
        Call fun806_RestaurarConfiguracion(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
        On Error GoTo 0
    End If
    
    ' Mostrar informaci�n detallada del error
    Call fun807_MostrarErrorDetallado(strFuncionActual, strTipoError, lngLineaError, Err.Number, Err.Description)
    
    F004_Forzar_Delimitadores_en_Excel = False
End Function


Public Function F004_Restaurar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCI�N PRINCIPAL: F004_Restaurar_Delimitadores_en_Excel
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura los delimitadores originales de Excel desde la hoja de respaldo
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables globales con valores por defecto (C2, C3, C4)
    ' 2. Obtener referencia al libro actual
    ' 3. Verificar si existe la hoja de delimitadores originales
    ' 4. Si no existe, crear la hoja y dejarla visible (situaci�n extra�a para restauraci�n)
    ' 5. Si existe, verificar su visibilidad y hacerla visible si est� oculta
    ' 6. Leer valores originales desde las celdas especificadas:
    '    - Use System Separators desde C2
    '    - Decimal Separator desde C3
    '    - Thousands Separator desde C4
    ' 7. Almacenar valores le�dos en variables globales correspondientes
    ' 8. Validar que los valores le�dos sean apropiados para restaurar
    ' 9. Aplicar configuraci�n original de delimitadores de Excel:
    '    - Use System Separators (True/False seg�n valor original)
    '    - Decimal Separator (car�cter seg�n valor original)
    '    - Thousands Separator (car�cter seg�n valor original)
    ' 10. Verificar variable global vOcultarRepostiorioDelimitadores
    ' 11. Si es True, ocultar la hoja de delimitadores al finalizar
    ' 12. Manejo exhaustivo de errores con informaci�n detallada y n�mero de l�nea
    '
    ' Par�metros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    ' Control de errores con n�mero de l�nea
    On Error GoTo ErrorHandler
    
    ' Variables locales
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim hojaExiste As Boolean
    Dim i As Integer
    Dim lineaError As Long
    Dim valorCelda As Variant
    
    ' Inicializar resultado como exitoso
    F004_Restaurar_Delimitadores_en_Excel = True
    
    ' ==========================================================================
    ' PASO 1: INICIALIZAR VARIABLES GLOBALES CON VALORES POR DEFECTO
    ' ==========================================================================
    lineaError = 100
    
    ' Variables para las celdas que contienen los valores originales
    ' NOTA: Usuario especific� C2 para todas, corrijo para C2, C3, C4 seg�n l�gica
    vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal = "C2"
    vCelda_Valor_Excel_DecimalSeparator_ValorOriginal = "C3"
    vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal = "C4"
    
    ' Variables para almacenar los valores originales (inicialmente vac�as)
    vExcel_UseSystemSeparators_ValorOriginal = ""
    vExcel_DecimalSeparator_ValorOriginal = ""
    vExcel_ThousandsSeparator_ValorOriginal = ""
    
    ' Usar la variable global ya definida para el nombre de la hoja
    If vHojaDelimitadoresExcelOriginales = "" Then
        vHojaDelimitadoresExcelOriginales = "06_Delimitadores_Originales"
    End If
    
    lineaError = 110
    
    ' ==========================================================================
    ' PASO 2: OBTENER REFERENCIA AL LIBRO ACTUAL
    ' ==========================================================================
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If wb Is Nothing Then
        F004_Restaurar_Delimitadores_en_Excel = False
        Exit Function
    End If
    
    lineaError = 120
    
    ' ==========================================================================
    ' PASO 3: VERIFICAR SI EXISTE LA HOJA DE DELIMITADORES ORIGINALES
    ' ==========================================================================
    
    hojaExiste = fun801_VerificarExistenciaHoja(wb, vHojaDelimitadoresExcelOriginales)
    
    lineaError = 130
    
    ' ==========================================================================
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEG�N CORRESPONDA
    ' ==========================================================================
    
    If Not hojaExiste Then
        ' La hoja no existe, crearla y dejarla visible
        ' NOTA: En un escenario de restauraci�n, esto ser�a extra�o, pero cumplimos la especificaci�n
        Set ws = fun802_CrearHojaDelimitadores(wb, vHojaDelimitadoresExcelOriginales)
        If ws Is Nothing Then
            F004_Restaurar_Delimitadores_en_Excel = False
            Exit Function
        End If
        ' Como no hay datos que leer, salir con �xito pero sin restaurar
        Debug.Print "ADVERTENCIA: Hoja de delimitadores creada, pero no hay valores para restaurar - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        Exit Function
    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
        Set ws = wb.Worksheets(vHojaDelimitadoresExcelOriginales)
        
        ' Verificar si est� oculta y hacerla visible si es necesario
        If Not fun803_HacerHojaVisible(ws) Then
            Debug.Print "ADVERTENCIA: No se pudo hacer visible la hoja " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        End If
    End If
    
    lineaError = 140
    
    ' ==========================================================================
    ' PASO 5: LEER VALORES ORIGINALES DESDE LAS CELDAS ESPECIFICADAS
    ' ==========================================================================
    
    ' Leer valor de Use System Separators desde C2
    valorCelda = ws.Range(vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal).Value
    vExcel_UseSystemSeparators_ValorOriginal = fun804_ConvertirValorACadena(valorCelda)
    
    ' Leer valor de Decimal Separator desde C3
    valorCelda = ws.Range(vCelda_Valor_Excel_DecimalSeparator_ValorOriginal).Value
    vExcel_DecimalSeparator_ValorOriginal = fun804_ConvertirValorACadena(valorCelda)
    
    ' Leer valor de Thousands Separator desde C4
    valorCelda = ws.Range(vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal).Value
    vExcel_ThousandsSeparator_ValorOriginal = fun804_ConvertirValorACadena(valorCelda)
    
    lineaError = 150
    
    ' ==========================================================================
    ' PASO 6: VALIDAR QUE SE HAYAN LE�DO VALORES V�LIDOS
    ' ==========================================================================
    
    If Not fun805_ValidarValoresOriginales() Then
        Debug.Print "ADVERTENCIA: No se encontraron valores v�lidos para restaurar en la hoja: " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        F004_Restaurar_Delimitadores_en_Excel = False
        Exit Function
    End If
    
    lineaError = 160
    
    ' ==========================================================================
    ' PASO 7: APLICAR CONFIGURACI�N ORIGINAL DE DELIMITADORES DE EXCEL
    ' ==========================================================================
    
    ' Restaurar Use System Separators (True/False)
    If Not fun806_RestaurarUseSystemSeparators(vExcel_UseSystemSeparators_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Use System Separators - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    ' Restaurar Decimal Separator (car�cter)
    If Not fun807_RestaurarDecimalSeparator(vExcel_DecimalSeparator_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Decimal Separator - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    ' Restaurar Thousands Separator (car�cter)
    If Not fun808_RestaurarThousandsSeparator(vExcel_ThousandsSeparator_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Thousands Separator - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    lineaError = 170
    
    ' ==========================================================================
    ' PASO 8: VERIFICAR SI DEBE OCULTAR LA HOJA
    ' ==========================================================================
    
    ' Verificar la variable global vOcultarRepostiorioDelimitadores
    If vOcultarRepostiorioDelimitadores = True Then
        ' Ocultar la hoja de delimitadores
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        End If
    End If
    
    lineaError = 180
    
    ' ==========================================================================
    ' PASO 9: FINALIZACI�N EXITOSA
    ' ==========================================================================
    
    Debug.Print "�XITO: Delimitadores restaurados correctamente - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Restaurar_Delimitadores_en_Excel = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: F004_Restaurar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now() & vbCrLf & _
                   "USUARIO: david-joaquin-corredera-de-colsa"
    
    ' Log del error para debugging
    Debug.Print mensajeError
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
    
End Function

