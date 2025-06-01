Attribute VB_Name = "Modulo_521_FUNC_Auxiliares"
Option Explicit


Public Function fun801_LogMessage(ByVal strMessage As String, _
                                Optional ByVal blnIsError As Boolean = False, _
                                Optional ByVal strFileName As String = "", _
                                Optional ByVal strSheetName As String = "") As Boolean
        
    '------------------------------------------------------------------------------
    ' FUNCI�N: fun801_LogMessage
    ' PROP�SITO: Sistema integral de logging para registrar eventos y errores
    '
    ' PAR�METROS:
    ' - strMessage (String): Mensaje a registrar
    ' - blnIsError (Boolean, Opcional): True=ERROR, False=INFO (defecto: False)
    ' - strFileName (String, Opcional): Archivo relacionado (defecto: "NA")
    ' - strSheetName (String, Opcional): Hoja relacionada (defecto: "NA")
    '
    ' RETORNA: Boolean - True si exitoso, False si error
    '
    ' FUNCIONALIDADES:
    ' - Crea hoja de log autom�ticamente con formato profesional
    ' - Timestamp ISO, usuario del sistema, tipo de evento
    ' - Formato condicional para errores (fondo rojo)
    ' - Filtros autom�ticos y ajuste de columnas
    '
    ' COMPATIBILIDAD: Excel 97-365, Office Online, SharePoint, Teams
    '
    ' EJEMPLO: Call fun801_LogMessage("Operaci�n completada", False, "datos.csv")
    '------------------------------------------------------------------------------
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para el log
    Dim wsLog As Worksheet
    Dim lngLastRow As Long
    Dim strDateTime As String
    Dim strUserName As String
    Dim strLogType As String
    
    ' Inicializaci�n
    strFuncion = "fun801_LogMessage"
    fun801_LogMessage = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Verificar hoja de log
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If Not fun802_SheetExists(gstrHoja_Log) Then
        If Not F002_Crear_Hoja(gstrHoja_Log) Then
            MsgBox "Error al crear la hoja de log", vbCritical
            Exit Function
        End If
        
        ' Crear y formatear encabezados
        With ThisWorkbook.Sheets(gstrHoja_Log)
            ' Establecer textos de encabezados exactamente como se solicita
            .Range("A1").Value = "Date/Time"
            .Range("B1").Value = "User"
            .Range("C1").Value = "Type"
            .Range("D1").Value = "File"
            .Range("E1").Value = "Sheet"
            .Range("F1").Value = "Message"
            
            ' Formato de encabezados
            With .Range("A1:F1")
                .Font.Bold = True
                .Font.Size = 11
                .Font.Name = "Calibri"
                .Interior.Color = RGB(200, 200, 200)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' Formato espec�fico para la columna de fecha
            .Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
            
            ' Ajustar anchos de columna
            .Columns("A").ColumnWidth = 20  ' Date/Time
            .Columns("B").ColumnWidth = 15  ' User
            .Columns("C").ColumnWidth = 15  ' Type
            .Columns("D").ColumnWidth = 40  ' File
            .Columns("E").ColumnWidth = 20  ' Sheet
            .Columns("F").ColumnWidth = 60  ' Message
            
            ' Filtros autom�ticos
            .Range("A1:F1").AutoFilter
        End With
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Preparar datos para el log
    '--------------------------------------------------------------------------
    lngLineaError = 55
    Set wsLog = ThisWorkbook.Sheets(gstrHoja_Log)
    
    ' Obtener �ltima fila
    lngLastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Preparar datos (reemplazar valores vac�os con "NA")
    strDateTime = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    strUserName = IIf(Environ("USERNAME") = "", "NA", Environ("USERNAME"))
    strLogType = IIf(blnIsError, "ERROR", "INFO")
    strFileName = IIf(Len(Trim(strFileName)) = 0, "NA", strFileName)
    strSheetName = IIf(Len(Trim(strSheetName)) = 0, "NA", strSheetName)
    strMessage = IIf(Len(Trim(strMessage)) = 0, "NA", strMessage)
    
    '--------------------------------------------------------------------------
    ' 3. Escribir en el log
    '--------------------------------------------------------------------------
    lngLineaError = 70
    With wsLog
        ' Escribir datos
        .Cells(lngLastRow, 1).Value = strDateTime    ' Date/Time
        .Cells(lngLastRow, 2).Value = strUserName    ' User
        .Cells(lngLastRow, 3).Value = strLogType     ' Type
        .Cells(lngLastRow, 4).Value = strFileName    ' File
        .Cells(lngLastRow, 5).Value = strSheetName   ' Sheet
        .Cells(lngLastRow, 6).Value = strMessage     ' Message
        
        ' Formato de la nueva fila
        With .Range(.Cells(lngLastRow, 1), .Cells(lngLastRow, 6))
            ' Formato general
            .Font.Name = "Calibri"
            .Font.Size = 10
            .VerticalAlignment = xlTop
            .WrapText = True
            
            ' Bordes
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            
            ' Formato condicional para errores
            If blnIsError Then
                .Interior.Color = RGB(255, 200, 200)
                .Font.Bold = True
            End If
        End With
        
        ' Asegurar formato de fecha en la columna A
        .Cells(lngLastRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    
    fun801_LogMessage = True
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    MsgBox strMensajeError, vbCritical, "Error en sistema de logging"
    fun801_LogMessage = False
End Function


Public Function fun802_SheetExists(ByVal strSheetName As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun802_SheetExists
    ' PROP�SITO:
    ' Verifica de forma segura si una hoja de c�lculo (worksheet) existe en el libro
    ' de Excel actual antes de intentar trabajar con ella.
    '******************************************************************************
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(strSheetName)
    On Error GoTo 0
    
    fun802_SheetExists = Not ws Is Nothing
End Function


Public Function F002_Crear_Hoja(ByVal strNombreHoja As String) As Boolean

    '******************************************************************************
    ' M�dulo: F002_Crear_Hoja
    ' Fecha y Hora de Creaci�n: 2025-05-26 09:17:15 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Funci�n para crear hojas en el libro con formato y configuraci�n est�ndar
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para manejo de hojas
    Dim ws As Worksheet
    
    ' Inicializaci�n
    strFuncion = "F002_Crear_Hoja"
    F002_Crear_Hoja = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Verificar si la hoja ya existe
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If fun802_SheetExists(strNombreHoja) Then
        F002_Crear_Hoja = True
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Crear nueva hoja
    '--------------------------------------------------------------------------
    lngLineaError = 40
    Application.ScreenUpdating = False
    
    ' Crear hoja al final del libro
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    ' Asignar nombre
    ws.Name = strNombreHoja
    
    ' Configuraci�n b�sica
    'With ws
    '    ' Ajustar vista
    '    .DisplayGridlines = True
    '    .DisplayHeadings = True
    '
    '    ' Configurar primera vista
    '    .Range("A1").Select
    '
    '    ' Ajustar ancho de columnas est�ndar
    '    .Columns.StandardWidth = 10
    '
    '    ' Configurar �rea de impresi�n
    '    .PageSetup.PrintArea = ""
    'End With
    
    Application.ScreenUpdating = True
    
    F002_Crear_Hoja = True
    Exit Function

GestorErrores:
    ' Restaurar configuraci�n
    Application.ScreenUpdating = True
    
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F002_Crear_Hoja = False
End Function



Public Function fun801_LimpiarHoja(ByVal strNombreHoja As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun801_LimpiarHoja
    ' FECHA Y HORA DE CREACI�N: 2025-05-28 17:50:26 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa    '
    ' PROP�SITO:
    ' Limpia de forma segura y eficiente todo el contenido de una hoja de c�lculo
    ' espec�fica, preservando el formato y estructura, pero eliminando todos los
    ' datos y valores almacenados en las celdas utilizadas.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(strNombreHoja)
    
    Application.ScreenUpdating = False
    ws.UsedRange.ClearContents
    Application.ScreenUpdating = True
    
    fun801_LimpiarHoja = True
    Exit Function
    
GestorErrores:
    fun801_LimpiarHoja = False
End Function

Public Function fun802_SeleccionarArchivo(ByVal strPrompt As String) As String
    
    '******************************************************************************
    ' FUNCI�N: fun802_SeleccionarArchivo (VERSI�N MEJORADA)
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' FECHA MODIFICACI�N: 2025-06-01
    '
    ' PROP�SITO:
    ' Proporciona una interfaz de usuario intuitiva para seleccionar archivos de
    ' texto (TXT y CSV) con sistema de carpetas de respaldo autom�tico.
    '
    ' L�GICA DE CARPETAS DE RESPALDO:
    ' 1. Carpeta del archivo Excel actual
    ' 2. %TEMP% (si hay error)
    ' 3. %TMP% (si hay error)
    ' 4. %USERPROFILE% (si hay error)
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para carpetas de respaldo
    Dim strCarpetaInicial As String
    Dim strCarpetaActual As String
    Dim intIntentoActual As Integer
    Dim blnCarpetaValida As Boolean
    
    ' Inicializaci�n
    strFuncion = "fun802_SeleccionarArchivo"
    fun802_SeleccionarArchivo = ""
    lngLineaError = 0
    intIntentoActual = 1
    blnCarpetaValida = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Intentar obtener carpetas de respaldo en orden de prioridad
    '--------------------------------------------------------------------------
    Do While intIntentoActual <= 4 And Not blnCarpetaValida
        lngLineaError = 40 + intIntentoActual
        
        Select Case intIntentoActual
            Case 1: ' Carpeta del archivo Excel actual
                strCarpetaActual = fun803_ObtenerCarpetaExcelActual()
                
            Case 2: ' Variable de entorno %TEMP%
                strCarpetaActual = fun804_ObtenerCarpetaTemp()
                
            Case 3: ' Variable de entorno %TMP%
                strCarpetaActual = fun805_ObtenerCarpetaTmp()
                
            Case 4: ' Variable de entorno %USERPROFILE%
                strCarpetaActual = fun806_ObtenerCarpetaUserProfile()
        End Select
        
        ' Verificar si la carpeta es v�lida y accesible
        If fun807_ValidarCarpeta(strCarpetaActual) Then
            blnCarpetaValida = True
            strCarpetaInicial = strCarpetaActual
        Else
            intIntentoActual = intIntentoActual + 1
        End If
    Loop
    
    ' Si no se pudo obtener ninguna carpeta v�lida, usar carpeta por defecto
    If Not blnCarpetaValida Then
        strCarpetaInicial = ""
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Mostrar di�logo de selecci�n de archivo
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    On Error GoTo GestorErrores
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = strPrompt
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt;*.csv"
        .AllowMultiSelect = False
        
        ' Establecer carpeta inicial si es v�lida
        If Len(strCarpetaInicial) > 0 Then
            .InitialFileName = strCarpetaInicial & "\"
        End If
        
        If .Show = -1 Then
            fun802_SeleccionarArchivo = .SelectedItems(1)
        Else
            fun802_SeleccionarArchivo = ""
        End If
    End With
    
    Exit Function
    
GestorErrores:
    ' Log del error
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Intento actual: " & intIntentoActual
    
    fun801_LogMessage strMensajeError, True
    fun802_SeleccionarArchivo = ""
End Function

Public Function fun803_ImportarArchivo(ByRef wsDestino As Worksheet, _
                                     ByVal strFilePath As String, _
                                     ByVal strColumnaInicial As String, _
                                     ByVal lngFilaInicial As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun803_ImportarArchivo
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Importa el contenido completo de archivos de texto plano (TXT/CSV) l�nea por
    ' l�nea hacia una hoja de Excel espec�fica, colocando cada l�nea del archivo
    ' en una celda individual seg�n la posici�n inicial definida. Funci�n core
    ' del sistema de importaci�n de datos de presupuesto.
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim objFSO As Object
    Dim objFile As Object
    Dim strLine As String
    Dim lngRow As Long
    
    ' Inicializaci�n
    strFuncion = "fun803_ImportarArchivo"
    fun803_ImportarArchivo = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros
    '--------------------------------------------------------------------------
    lngLineaError = 20
    If wsDestino Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, "Hoja de destino no v�lida"
    End If
    
    If Len(strFilePath) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, "Ruta de archivo no v�lida"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Configurar objetos para lectura de archivo
    '--------------------------------------------------------------------------
    lngLineaError = 35
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFilePath, 1) ' ForReading = 1
    
    '--------------------------------------------------------------------------
    ' 3. Leer archivo l�nea por l�nea
    '--------------------------------------------------------------------------
    lngLineaError = 45
    lngRow = lngFilaInicial
    
    While Not objFile.AtEndOfStream
        strLine = objFile.ReadLine
        wsDestino.Range(strColumnaInicial & lngRow).Value = strLine
        lngRow = lngRow + 1
    Wend
    
    '--------------------------------------------------------------------------
    ' 4. Limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 60
    objFile.Close
    Set objFile = Nothing
    Set objFSO = Nothing
    
    fun803_ImportarArchivo = True
    Exit Function

GestorErrores:
    ' Limpieza en caso de error
    If Not objFile Is Nothing Then
        objFile.Close
        Set objFile = Nothing
    End If
    Set objFSO = Nothing
    
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun803_ImportarArchivo = False
End Function


Public Function fun804_DetectarRangoDatos(ByRef ws As Worksheet, _
                                         ByRef lngLineaInicial As Long, _
                                         ByRef lngLineaFinal As Long) As Boolean
    '******************************************************************************
    ' FUNCI�N: fun804_DetectarRangoDatos
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Detecta autom�ticamente el rango exacto de datos en una columna espec�fica
    ' de una hoja de c�lculo, identificando la primera y �ltima fila que contienen
    ' informaci�n. Funci�n esencial para determinar l�mites de procesamiento
    ' despu�s de la importaci�n de datos.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim rngBusqueda As Range
    Dim lngColumna As Long
    
    ' Obtener n�mero de columna
    lngColumna = Range(vColumnaInicial_Importacion & "1").Column
    
    ' Configurar rango de b�squeda
    Set rngBusqueda = ws.Columns(lngColumna)
    
    With rngBusqueda
        ' Encontrar primera celda con datos
        Set rngBusqueda = .Find(What:="*", _
                               After:=.Cells(.Cells.Count), _
                               LookIn:=xlFormulas, _
                               LookAt:=xlPart, _
                               SearchOrder:=xlByRows, _
                               SearchDirection:=xlNext)
        
        If Not rngBusqueda Is Nothing Then
            lngLineaInicial = rngBusqueda.Row
            
            ' Encontrar �ltima celda con datos
            Set rngBusqueda = .Find(What:="*", _
                                   After:=.Cells(1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlPrevious)
            
            lngLineaFinal = rngBusqueda.Row
            fun804_DetectarRangoDatos = True
        Else
            lngLineaInicial = 0
            lngLineaFinal = 0
            fun804_DetectarRangoDatos = False
        End If
    End With
    Exit Function
    
GestorErrores:
    lngLineaInicial = 0
    lngLineaFinal = 0
    fun804_DetectarRangoDatos = False
End Function




Public Function fun801_VerificarExistenciaHoja(wb As Workbook, nombreHoja As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 801: VERIFICAR EXISTENCIA DE HOJA
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Verifica si una hoja existe en el libro especificado
    ' Par�metros: wb (Workbook), nombreHoja (String)
    ' Retorna: Boolean (True si existe, False si no existe)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim lineaError As Long
    
    lineaError = 200
    fun801_VerificarExistenciaHoja = False
    
    ' Verificar par�metros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Exit Function
    End If
    
    lineaError = 210
    
    ' Recorrer todas las hojas del libro (m�todo compatible con Excel 97)
    For i = 1 To wb.Worksheets.Count
        If UCase(wb.Worksheets(i).Name) = UCase(nombreHoja) Then
            fun801_VerificarExistenciaHoja = True
            Exit For
        End If
    Next i
    
    lineaError = 220
    
    Exit Function
    
ErrorHandler:
    fun801_VerificarExistenciaHoja = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun801_VerificarExistenciaHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function



Public Sub fun804_LimpiarContenidoHoja(ws As Worksheet)
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 804: LIMPIAR CONTENIDO DE HOJA
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Limpia todo el contenido de una hoja espec�fica
    ' Par�metros: ws (Worksheet)
    ' Retorna: Nada (Sub procedure)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 500
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        Exit Sub
    End If
    
    lineaError = 510
    
    ' Verificar que la hoja no est� protegida
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    lineaError = 520
    
    ' Limpiar todo el contenido de la hoja (m�todo compatible con todas las versiones)
    ws.Cells.Clear
    
    lineaError = 530
    
    Exit Sub
    
ErrorHandler:
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun804_LimpiarContenidoHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub


Public Function fun805_DetectarUseSystemSeparators() As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 805: DETECTAR USE SYSTEM SEPARATORS
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta si Excel est� usando separadores del sistema
    ' Par�metros: Ninguno
    ' Retorna: String ("True" o "False")
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variable para almacenar el resultado
    Dim resultado As String
    Dim lineaError As Long
    
    lineaError = 600
    
    ' Detectar configuraci�n actual de Use System Separators
    ' Usar compilaci�n condicional para compatibilidad con versiones
    
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lineaError = 610
        If Application.UseSystemSeparators Then
            resultado = "True"
        Else
            resultado = "False"
        End If
    #Else
        ' Excel 97, 2003 y anteriores
        lineaError = 620
        resultado = fun809_DetectarUseSystemSeparatorsLegacy()
    #End If
    
    lineaError = 630
    
    fun805_DetectarUseSystemSeparators = resultado
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, intentar m�todo alternativo
    fun805_DetectarUseSystemSeparators = fun809_DetectarUseSystemSeparatorsLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun805_DetectarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun806_DetectarDecimalSeparator() As String

    ' =============================================================================
    ' FUNCI�N AUXILIAR 806: DETECTAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta el separador decimal actual de Excel
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador decimal)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    
    ' Detectar separador decimal actual (compatible con todas las versiones)
    fun806_DetectarDecimalSeparator = Application.DecimalSeparator
    
    lineaError = 710
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar m�todo alternativo
    fun806_DetectarDecimalSeparator = fun810_DetectarDecimalSeparatorLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun806_DetectarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun807_DetectarThousandsSeparator() As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 807: DETECTAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta el separador de miles actual de Excel
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador de miles)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    
    ' Detectar separador de miles actual (compatible con todas las versiones)
    fun807_DetectarThousandsSeparator = Application.ThousandsSeparator
    
    lineaError = 810
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar m�todo alternativo
    fun807_DetectarThousandsSeparator = fun811_DetectarThousandsSeparatorLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun807_DetectarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Sub fun808_OcultarHojaDelimitadores(ws As Worksheet)

    ' =============================================================================
    ' FUNCI�N AUXILIAR 808: OCULTAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Oculta la hoja de delimitadores si est� habilitada la opci�n
    ' Par�metros: ws (Worksheet)
    ' Retorna: Nada (Sub procedure)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 900
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        Exit Sub
    End If
    
    lineaError = 910
    
    ' Verificar que el libro permite ocultar hojas (no protegido)
    If ws.Parent.ProtectStructure Then
        Exit Sub
    End If
    
    lineaError = 920
    
    ' Ocultar la hoja (compatible con todas las versiones de Excel)
    ws.Visible = xlSheetHidden
    
    lineaError = 930
    
    Exit Sub
    
ErrorHandler:
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun808_OcultarHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub


Public Function fun809_DetectarUseSystemSeparatorsLegacy() As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 809: DETECTAR USE SYSTEM SEPARATORS (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar Use System Separators en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String ("True" o "False")
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para comparaci�n
    Dim separadorSistema As String
    Dim separadorExcel As String
    Dim lineaError As Long
    
    lineaError = 1000
    
    ' Obtener separador decimal del sistema (Windows)
    ' M�todo compatible con Excel 97 y 2003
    separadorSistema = Mid(CStr(1.1), 2, 1)
    
    lineaError = 1010
    
    ' Obtener separador decimal de Excel
    separadorExcel = Application.DecimalSeparator
    
    lineaError = 1020
    
    ' Si coinciden, probablemente Use System Separators est� activado
    If separadorSistema = separadorExcel Then
        fun809_DetectarUseSystemSeparatorsLegacy = "True"
    Else
        fun809_DetectarUseSystemSeparatorsLegacy = "False"
    End If
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir False por defecto
    fun809_DetectarUseSystemSeparatorsLegacy = "False"
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun809_DetectarUseSystemSeparatorsLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun810_DetectarDecimalSeparatorLegacy() As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 810: DETECTAR DECIMAL SEPARATOR (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar separador decimal en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador decimal)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detecci�n
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1100
    
    ' M�todo alternativo: formatear un n�mero y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = CStr(1.1)
    
    lineaError = 1110
    
    ' El separador decimal es el segundo car�cter en el formato est�ndar
    If Len(numeroFormateado) >= 2 Then
        fun810_DetectarDecimalSeparatorLegacy = Mid(numeroFormateado, 2, 1)
    Else
        ' Fallback: asumir punto por defecto
        fun810_DetectarDecimalSeparatorLegacy = "."
    End If
    
    lineaError = 1120
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir punto por defecto
    fun810_DetectarDecimalSeparatorLegacy = "."
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun810_DetectarDecimalSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun811_DetectarThousandsSeparatorLegacy() As String

    ' =============================================================================
    ' FUNCI�N AUXILIAR 811: DETECTAR THOUSANDS SEPARATOR (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar separador de miles en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador de miles)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detecci�n
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1200
    
    ' M�todo alternativo: formatear un n�mero grande y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = Format(1000, "#,##0")
    
    lineaError = 1210
    
    ' El separador de miles es el segundo car�cter en n�meros de 4 d�gitos
    If Len(numeroFormateado) >= 2 Then
        fun811_DetectarThousandsSeparatorLegacy = Mid(numeroFormateado, 2, 1)
    Else
        ' Si no hay separador visible, asumir coma por defecto
        fun811_DetectarThousandsSeparatorLegacy = ","
    End If
    
    lineaError = 1220
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir coma por defecto
    fun811_DetectarThousandsSeparatorLegacy = ","
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun811_DetectarThousandsSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function



Public Function fun802_CrearHojaDelimitadores(wb As Workbook, nombreHoja As String) As Worksheet

    ' =============================================================================
    ' FUNCI�N AUXILIAR 802: CREAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Crea una nueva hoja con el nombre especificado y la deja visible
    ' Par�metros: wb (Workbook), nombreHoja (String)
    ' Retorna: Worksheet (referencia a la hoja creada, Nothing si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lineaError As Long
    
    lineaError = 300
    
    ' Verificar par�metros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Exit Function
    End If
    
    lineaError = 310
    
    ' Verificar que el libro no est� protegido (importante para entornos cloud)
    If wb.ProtectStructure Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Debug.Print "ERROR: No se puede crear hoja, libro protegido - Funci�n: fun802_CrearHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 320
    
    ' Crear nueva hoja al final del libro (m�todo compatible con todas las versiones)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    lineaError = 330
    
    ' Asignar nombre a la hoja
    ws.Name = nombreHoja
    
    lineaError = 340
    
    ' Asegurar que la hoja est� visible
    ws.Visible = xlSheetVisible
    
    lineaError = 350
    
    ' Configuraci�n adicional para compatibilidad con entornos cloud
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    ' Retornar referencia a la hoja creada
    Set fun802_CrearHojaDelimitadores = ws
    
    lineaError = 360
    
    Exit Function
    
ErrorHandler:
    Set fun802_CrearHojaDelimitadores = Nothing
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun802_CrearHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun803_HacerHojaVisible(ws As Worksheet) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 803: HACER HOJA VISIBLE
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Verifica la visibilidad de una hoja y la hace visible si est� oculta
    ' Par�metros: ws (Worksheet)
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 400
    fun803_HacerHojaVisible = True
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        fun803_HacerHojaVisible = False
        Exit Function
    End If
    
    lineaError = 410
    
    ' Verificar que el libro permite cambiar visibilidad (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede cambiar visibilidad, libro protegido - Funci�n: fun803_HacerHojaVisible - " & Now()
        Exit Function
    End If
    
    lineaError = 420
    
    ' Verificar el estado actual de visibilidad y actuar seg�n corresponda
    Select Case ws.Visible
        Case xlSheetVisible
            ' La hoja ya est� visible, no hacer nada
            Debug.Print "INFO: Hoja " & ws.Name & " ya est� visible - Funci�n: fun803_HacerHojaVisible - " & Now()
            
        Case xlSheetHidden, xlSheetVeryHidden
            ' La hoja est� oculta, hacerla visible
            ws.Visible = xlSheetVisible
            Debug.Print "INFO: Hoja " & ws.Name & " se hizo visible - Funci�n: fun803_HacerHojaVisible - " & Now()
            
        Case Else
            ' Estado desconocido, forzar visibilidad
            ws.Visible = xlSheetVisible
            Debug.Print "INFO: Hoja " & ws.Name & " visibilidad forzada - Funci�n: fun803_HacerHojaVisible - " & Now()
    End Select
    
    lineaError = 430
    
    Exit Function
    
ErrorHandler:
    fun803_HacerHojaVisible = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun803_HacerHojaVisible" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun804_ConvertirValorACadena(valor As Variant) As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 804: CONVERTIR VALOR A CADENA
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Convierte un valor de celda a cadena de texto de forma segura
    ' Par�metros: valor (Variant)
    ' Retorna: String (valor convertido o cadena vac�a si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    Dim resultado As String
    
    lineaError = 500
    
    ' Verificar si el valor es Nothing o Empty
    If IsEmpty(valor) Or IsNull(valor) Then
        resultado = ""
    ElseIf IsError(valor) Then
        resultado = ""
    Else
        ' Convertir a cadena
        resultado = CStr(valor)
        ' Eliminar espacios en blanco al inicio y final
        resultado = Trim(resultado)
    End If
    
    lineaError = 510
    
    fun804_ConvertirValorACadena = resultado
    
    Exit Function
    
ErrorHandler:
    fun804_ConvertirValorACadena = ""
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun804_ConvertirValorACadena" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun805_ValidarValoresOriginales() As Boolean

    ' =============================================================================
    ' FUNCI�N AUXILIAR 805: VALIDAR VALORES ORIGINALES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Valida que los valores originales le�dos sean v�lidos para restaurar
    ' Par�metros: Ninguno (usa variables globales)
    ' Retorna: Boolean (True si v�lidos, False si no v�lidos)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    Dim esValido As Boolean
    
    lineaError = 600
    esValido = True
    
    ' Validar Use System Separators (debe ser "True" o "False")
    If vExcel_UseSystemSeparators_ValorOriginal <> "True" And vExcel_UseSystemSeparators_ValorOriginal <> "False" Then
        If vExcel_UseSystemSeparators_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inv�lido para Use System Separators: '" & vExcel_UseSystemSeparators_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 610
    
    ' Validar Decimal Separator (debe ser un solo car�cter)
    If Len(vExcel_DecimalSeparator_ValorOriginal) <> 1 Then
        If vExcel_DecimalSeparator_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inv�lido para Decimal Separator: '" & vExcel_DecimalSeparator_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 620
    
    ' Validar Thousands Separator (debe ser un solo car�cter)
    If Len(vExcel_ThousandsSeparator_ValorOriginal) <> 1 Then
        If vExcel_ThousandsSeparator_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inv�lido para Thousands Separator: '" & vExcel_ThousandsSeparator_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 630
    
    fun805_ValidarValoresOriginales = esValido
    
    ' Log de valores validados
    If esValido Then
        Debug.Print "INFO: Valores v�lidos para restaurar - UseSystem:" & vExcel_UseSystemSeparators_ValorOriginal & " Decimal:'" & vExcel_DecimalSeparator_ValorOriginal & "' Thousands:'" & vExcel_ThousandsSeparator_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
    End If
    
    Exit Function
    
ErrorHandler:
    fun805_ValidarValoresOriginales = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun805_ValidarValoresOriginales" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun806_RestaurarUseSystemSeparators(valorOriginal As String) As Boolean

    ' =============================================================================
    ' FUNCI�N AUXILIAR 806: RESTAURAR USE SYSTEM SEPARATORS
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura la configuraci�n de Use System Separators
    ' Par�metros: valorOriginal (String) - "True" o "False"
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    fun806_RestaurarUseSystemSeparators = True
    
    ' Verificar que el valor sea v�lido
    If valorOriginal <> "True" And valorOriginal <> "False" Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Use System Separators, valor inv�lido: '" & valorOriginal & "' - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        fun806_RestaurarUseSystemSeparators = False
        Exit Function
    End If
    
    lineaError = 710
    
    ' Usar compilaci�n condicional para compatibilidad con versiones
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lineaError = 720
        If valorOriginal = "True" Then
            Application.UseSystemSeparators = True
            Debug.Print "INFO: Use System Separators configurado a True - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        Else
            Application.UseSystemSeparators = False
            Debug.Print "INFO: Use System Separators configurado a False - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        End If
    #Else
        ' Excel 97, 2003 y anteriores
        lineaError = 730
        Debug.Print "ADVERTENCIA: Use System Separators no disponible en esta versi�n de Excel - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        ' En versiones antiguas, esta propiedad no existe, pero no es error
    #End If
    
    lineaError = 740
    
    Exit Function
    
ErrorHandler:
    fun806_RestaurarUseSystemSeparators = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun806_RestaurarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun807_RestaurarDecimalSeparator(valorOriginal As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 807: RESTAURAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura el separador decimal original
    ' Par�metros: valorOriginal (String) - car�cter del separador decimal
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    fun807_RestaurarDecimalSeparator = True
    
    ' Verificar que el valor sea v�lido (un solo car�cter)
    If Len(valorOriginal) <> 1 Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Decimal Separator, valor inv�lido: '" & valorOriginal & "' - Funci�n: fun807_RestaurarDecimalSeparator - " & Now()
        fun807_RestaurarDecimalSeparator = False
        Exit Function
    End If
    
    lineaError = 810
    
    ' Restaurar separador decimal (compatible con todas las versiones)
    Application.DecimalSeparator = valorOriginal
    Debug.Print "INFO: Decimal Separator restaurado a: '" & valorOriginal & "' - Funci�n: fun807_RestaurarDecimalSeparator - " & Now()
    
    lineaError = 820
    
    Exit Function
    
ErrorHandler:
    fun807_RestaurarDecimalSeparator = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun807_RestaurarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun808_RestaurarThousandsSeparator(valorOriginal As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 808: RESTAURAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura el separador de miles original
    ' Par�metros: valorOriginal (String) - car�cter del separador de miles
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 900
    fun808_RestaurarThousandsSeparator = True
    
    ' Verificar que el valor sea v�lido (un solo car�cter)
    If Len(valorOriginal) <> 1 Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Thousands Separator, valor inv�lido: '" & valorOriginal & "' - Funci�n: fun808_RestaurarThousandsSeparator - " & Now()
        fun808_RestaurarThousandsSeparator = False
        Exit Function
    End If
    
    lineaError = 910
    
    ' Restaurar separador de miles (compatible con todas las versiones)
    Application.ThousandsSeparator = valorOriginal
    Debug.Print "INFO: Thousands Separator restaurado a: '" & valorOriginal & "' - Funci�n: fun808_RestaurarThousandsSeparator - " & Now()
    
    lineaError = 920
    
    Exit Function
    
ErrorHandler:
    fun808_RestaurarThousandsSeparator = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun808_RestaurarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCI�N AUXILIAR 809: OCULTAR HOJA DE DELIMITADORES
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripci�n: Oculta la hoja de delimitadores si est� habilitada la opci�n
' Par�metros: ws (Worksheet)
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Function fun809_OcultarHojaDelimitadores(ws As Worksheet) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 1000
    fun809_OcultarHojaDelimitadores = True
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        fun809_OcultarHojaDelimitadores = False
        Exit Function
    End If
    
    lineaError = 1010
    
    ' Verificar que el libro permite ocultar hojas (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede ocultar hoja, libro protegido - Funci�n: fun809_OcultarHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 1020
    
    ' Ocultar la hoja (compatible con todas las versiones de Excel)
    ws.Visible = xlSheetHidden
    Debug.Print "INFO: Hoja " & ws.Name & " ocultada - Funci�n: fun809_OcultarHojaDelimitadores - " & Now()
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    fun809_OcultarHojaDelimitadores = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun809_OcultarHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun802_VerificarCompatibilidad() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun802_VerificarCompatibilidad
    ' PROP�SITO: Verifica compatibilidad con diferentes versiones de Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = compatible, False = no compatible)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun802
    
    Dim strVersionExcel As String
    Dim dblVersionNumero As Double
    
    ' Obtener versi�n de Excel
    strVersionExcel = Application.Version
    dblVersionNumero = CDbl(strVersionExcel)
    
    ' Verificar compatibilidad (Excel 97 = 8.0, 2003 = 11.0, 365 = 16.0+)
    If dblVersionNumero >= 8# Then
        fun802_VerificarCompatibilidad = True
    Else
        fun802_VerificarCompatibilidad = False
    End If
    
    Exit Function

ErrorHandler_fun802:
    ' En caso de error, asumir compatibilidad
    fun802_VerificarCompatibilidad = True
End Function

Public Sub fun803_ObtenerConfiguracionActual(ByRef strDecimalAnterior As String, ByRef strMilesAnterior As String)
    ' =============================================================================
    ' FUNCI�N: fun803_ObtenerConfiguracionActual
    ' PROP�SITO: Obtiene la configuraci�n actual de delimitadores
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error GoTo ErrorHandler_fun803
    
    ' Obtener delimitador decimal actual
    strDecimalAnterior = Application.International(xlDecimalSeparator)
    
    ' Obtener delimitador de miles actual
    strMilesAnterior = Application.International(xlThousandsSeparator)
    
    Exit Sub

ErrorHandler_fun803:
    ' En caso de error, usar valores por defecto
    strDecimalAnterior = "."
    strMilesAnterior = ","
End Sub

Public Function fun804_AplicarNuevosDelimitadores() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun804_AplicarNuevosDelimitadores
    ' PROP�SITO: Aplica los nuevos delimitadores al sistema
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = �xito, False = error)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun804
    
    ' Aplicar nuevo delimitador decimal
    Application.DecimalSeparator = vDelimitadorDecimal_HFM
    
    ' Aplicar nuevo delimitador de miles
    Application.ThousandsSeparator = vDelimitadorMiles_HFM
    
    ' Forzar que Excel use los delimitadores del sistema
    Application.UseSystemSeparators = False
    
    ' Actualizar la pantalla
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    fun804_AplicarNuevosDelimitadores = True
    Exit Function

ErrorHandler_fun804:
    fun804_AplicarNuevosDelimitadores = False
End Function

Public Function fun805_VerificarAplicacionDelimitadores() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun805_VerificarAplicacionDelimitadores
    ' PROP�SITO: Verifica que los delimitadores se aplicaron correctamente
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = aplicados correctamente, False = error)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun805
    
    Dim strDecimalActual As String
    Dim strMilesActual As String
    
    ' Obtener delimitadores actuales
    strDecimalActual = Application.DecimalSeparator
    strMilesActual = Application.ThousandsSeparator
    
    ' Verificar que coinciden con los deseados
    If strDecimalActual = vDelimitadorDecimal_HFM And strMilesActual = vDelimitadorMiles_HFM Then
        fun805_VerificarAplicacionDelimitadores = True
    Else
        fun805_VerificarAplicacionDelimitadores = False
    End If
    
    Exit Function

ErrorHandler_fun805:
    fun805_VerificarAplicacionDelimitadores = False
End Function

Public Sub fun806_RestaurarConfiguracion(ByVal strDecimalAnterior As String, ByVal strMilesAnterior As String)
    ' =============================================================================
    ' FUNCI�N: fun806_RestaurarConfiguracion
    ' PROP�SITO: Restaura la configuraci�n anterior en caso de error
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error Resume Next
    
    ' Restaurar delimitador decimal anterior
    Application.DecimalSeparator = strDecimalAnterior
    
    ' Restaurar delimitador de miles anterior
    Application.ThousandsSeparator = strMilesAnterior
    
    ' Restaurar uso de separadores del sistema
    Application.UseSystemSeparators = True
    
    On Error GoTo 0
End Sub

Public Sub fun807_MostrarErrorDetallado(ByVal strFuncion As String, ByVal strTipoError As String, _
                                        ByVal lngLinea As Long, ByVal lngNumeroError As Long, _
                                        ByVal strDescripcionError As String)
    
    ' =============================================================================
    ' FUNCI�N: fun807_MostrarErrorDetallado
    ' PROP�SITO: Muestra informaci�n detallada del error ocurrido
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    Dim strMensajeError As String
    
    ' Construir mensaje de error detallado
    strMensajeError = "ERROR EN FUNCI�N DE DELIMITADORES" & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Funci�n: " & strFuncion & vbCrLf
    strMensajeError = strMensajeError & "Tipo de Error: " & strTipoError & vbCrLf
    strMensajeError = strMensajeError & "L�nea Aproximada: " & CStr(lngLinea) & vbCrLf
    strMensajeError = strMensajeError & "N�mero de Error VBA: " & CStr(lngNumeroError) & vbCrLf
    strMensajeError = strMensajeError & "Descripci�n: " & strDescripcionError & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Fecha/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Mostrar mensaje de error
    MsgBox strMensajeError, vbCritical, "Error en F004_Forzar_Delimitadores_en_Excel"
    
End Sub

' Funci�n auxiliar para obtener la primera fila vac�a despu�s del rango de datos
Public Function fun812_ObtenerPrimeraFilaVacia(ByRef ws As Worksheet, ByVal lngUltimaFilaDatos As Long) As Long
    '******************************************************************************
    ' FUNCI�N: fun812_ObtenerPrimeraFilaVacia
    ' FECHA Y HORA DE CREACI�N: 2025-05-30 08:15:21 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Localiza la primera fila completamente vac�a despu�s de un rango espec�fico de datos.
    ' Esta funci�n es cr�tica para encontrar la posici�n correcta donde insertar filas
    ' de resumen en el proceso de consolidaci�n de l�neas duplicadas.
    '
    ' PAR�METROS:
    ' - ws: Referencia a la hoja de c�lculo donde buscar
    ' - lngUltimaFilaDatos: N�mero de la �ltima fila del rango de datos actual
    '
    ' RETORNA:
    ' Long - N�mero de la primera fila vac�a encontrada
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngFilaActual As Long
    Dim lngColumna As Long
    Dim blnFilaVacia As Boolean
    
    ' Iniciar la b�squeda desde la fila siguiente a la �ltima con datos
    lngFilaActual = lngUltimaFilaDatos + 1
    
    ' Bucle para verificar filas hasta encontrar una vac�a
    Do
        blnFilaVacia = True  ' Asumir que la fila est� vac�a inicialmente
        
        ' Verificar si hay alguna celda con contenido en la fila
        For lngColumna = 1 To 50  ' Revisar las primeras 50 columnas (ajustar seg�n necesidad)
            If Len(Trim(CStr(ws.Cells(lngFilaActual, lngColumna).Value))) > 0 Then
                blnFilaVacia = False
                Exit For
            End If
        Next lngColumna
        
        ' Si encontramos una fila vac�a, devolver su n�mero
        If blnFilaVacia Then
            fun812_ObtenerPrimeraFilaVacia = lngFilaActual
            Exit Function
        End If
        
        ' Avanzar a la siguiente fila
        lngFilaActual = lngFilaActual + 1
        
    Loop While lngFilaActual <= ws.Rows.Count  ' Evitar bucle infinito
    
    ' Si llegamos aqu� sin encontrar fila vac�a, devolver un valor seguro
    fun812_ObtenerPrimeraFilaVacia = lngUltimaFilaDatos + 10
    Exit Function
    
GestorErrores:
    ' En caso de error, devolver un valor seguro
    fun812_ObtenerPrimeraFilaVacia = lngUltimaFilaDatos + 10
End Function


Public Function fun803_ObtenerCarpetaExcelActual() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta donde est� ubicado el archivo Excel actual
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener ruta completa del archivo actual
    If ThisWorkbook.Path <> "" Then
        strCarpeta = ThisWorkbook.Path
    ElseIf ActiveWorkbook.Path <> "" Then
        strCarpeta = ActiveWorkbook.Path
    Else
        strCarpeta = ""
    End If
    
    fun803_ObtenerCarpetaExcelActual = strCarpeta
    Exit Function
    
ErrorHandler:
    fun803_ObtenerCarpetaExcelActual = ""
End Function

Public Function fun804_ObtenerCarpetaTemp() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TEMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TEMP (compatible con Excel 97+)
    strCarpeta = Environ("TEMP")
    
    fun804_ObtenerCarpetaTemp = strCarpeta
    Exit Function
    
ErrorHandler:
    fun804_ObtenerCarpetaTemp = ""
End Function

Public Function fun805_ObtenerCarpetaTmp() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TMP (compatible con Excel 97+)
    strCarpeta = Environ("TMP")
    
    fun805_ObtenerCarpetaTmp = strCarpeta
    Exit Function
    
ErrorHandler:
    fun805_ObtenerCarpetaTmp = ""
End Function

Public Function fun806_ObtenerCarpetaUserProfile() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %USERPROFILE%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno USERPROFILE (compatible con Excel 97+)
    strCarpeta = Environ("USERPROFILE")
    
    fun806_ObtenerCarpetaUserProfile = strCarpeta
    Exit Function
    
ErrorHandler:
    fun806_ObtenerCarpetaUserProfile = ""
End Function

Public Function fun807_ValidarCarpeta(ByVal strCarpeta As String) As Boolean

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Valida si una carpeta existe y es accesible
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim objFSO As Object
    Dim blnResultado As Boolean
    
    blnResultado = False
    
    ' Verificar que la carpeta no est� vac�a
    If Len(Trim(strCarpeta)) = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Crear objeto FileSystemObject (compatible con Excel 97+)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si la carpeta existe y es accesible
    If objFSO.FolderExists(strCarpeta) Then
        blnResultado = True
    End If
    
    Set objFSO = Nothing
    fun807_ValidarCarpeta = blnResultado
    Exit Function
    
ErrorHandler:
    Set objFSO = Nothing
    fun807_ValidarCarpeta = False
End Function


'*******************************
'*******************************
Public Function fun812_CopiarContenidoCompleto(ByRef wsOrigen As Worksheet, _
                                               ByRef wsDestino As Worksheet) As Boolean
    
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR CORREGIDA: fun812_CopiarContenidoCompleto
    ' Fecha y Hora de Modificaci�n: 2025-06-01 19:34:00 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Copia todo el contenido de una hoja de trabajo a otra hoja de destino
    ' MANTENIENDO LA POSICI�N ORIGINAL de los datos (ej: si origen est� en B2,
    ' destino tambi�n estar� en B2).
    '******************************************************************************
    On Error GoTo GestorErrores
    
    Dim rngUsedOrigen As Range
    Dim strCeldaDestino As String
    
    ' Limpiar hoja destino
    If Not fun801_LimpiarHoja(wsDestino.Name) Then
        fun812_CopiarContenidoCompleto = False
        Exit Function
    End If
    
    ' Verificar que hay contenido en la hoja origen
    If wsOrigen.UsedRange Is Nothing Then
        fun812_CopiarContenidoCompleto = True
        Exit Function
    End If
    
    ' Obtener rango usado de origen
    Set rngUsedOrigen = wsOrigen.UsedRange
    
    ' Calcular celda destino manteniendo posici�n original
    ' Si el rango origen empieza en B2, el destino tambi�n empezar� en B2
    strCeldaDestino = wsDestino.Cells(rngUsedOrigen.Row, rngUsedOrigen.Column).Address
    
    ' Copiar manteniendo posici�n original
    rngUsedOrigen.Copy wsDestino.Range(strCeldaDestino)
    Application.CutCopyMode = False
    
    fun812_CopiarContenidoCompleto = True
    Exit Function
    
GestorErrores:
    Application.CutCopyMode = False
    fun812_CopiarContenidoCompleto = False
End Function


Public Function fun813_DetectarRangoCompleto(ByRef ws As Worksheet, _
                                            ByRef vFila_Inicial As Long, _
                                            ByRef vFila_Final As Long, _
                                            ByRef vColumna_Inicial As Long, _
                                            ByRef vColumna_Final As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun813_DetectarRangoCompleto
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim rngUsado As Range
    
    ' Obtener rango usado
    Set rngUsado = ws.UsedRange
    
    If rngUsado Is Nothing Then
        vFila_Inicial = 0
        vFila_Final = 0
        vColumna_Inicial = 0
        vColumna_Final = 0
        fun813_DetectarRangoCompleto = False
        Exit Function
    End If
    
    ' Detectar rangos
    vFila_Inicial = rngUsado.Row
    vFila_Final = rngUsado.Row + rngUsado.Rows.Count - 1
    vColumna_Inicial = rngUsado.Column
    vColumna_Final = rngUsado.Column + rngUsado.Columns.Count - 1
    
    fun813_DetectarRangoCompleto = True
    Exit Function
    
GestorErrores:
    vFila_Inicial = 0
    vFila_Final = 0
    vColumna_Inicial = 0
    vColumna_Final = 0
    fun813_DetectarRangoCompleto = False
End Function


Public Sub fun814_MostrarInformacionColumnas(ByVal vColumna_Inicial As Long, _
                                            ByVal vColumna_Final As Long, _
                                            ByVal vColumna_IdentificadorDeLinea As Long, _
                                            ByVal vColumna_LineaRepetida As Long, _
                                            ByVal vColumna_LineaTratada As Long, _
                                            ByVal vColumna_LineaSuma As Long, _
                                            ByVal vFila_Inicial As Long, _
                                            ByVal vFila_Final As Long)
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun814_MostrarInformacionColumnas
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    Dim strMensaje As String
    
    strMensaje = "INFORMACI�N DE VARIABLES DE COLUMNAS DE CONTROL" & vbCrLf & vbCrLf & _
                 "RANGOS DETECTADOS:" & vbCrLf & _
                 "- Fila Inicial: " & vFila_Inicial & vbCrLf & _
                 "- Fila Final: " & vFila_Final & vbCrLf & _
                 "- Columna Inicial: " & vColumna_Inicial & vbCrLf & _
                 "- Columna Final: " & vColumna_Final & vbCrLf & vbCrLf & _
                 "COLUMNAS DE CONTROL CALCULADAS:" & vbCrLf & _
                 "- vColumna_IdentificadorDeLinea = " & vColumna_IdentificadorDeLinea & _
                 " (Inicial+" & (vColumna_IdentificadorDeLinea - vColumna_Inicial) & ")" & vbCrLf & _
                 "- vColumna_LineaRepetida = " & vColumna_LineaRepetida & _
                 " (Inicial+" & (vColumna_LineaRepetida - vColumna_Inicial) & ")" & vbCrLf & _
                 "- vColumna_LineaTratada = " & vColumna_LineaTratada & _
                 " (Inicial+" & (vColumna_LineaTratada - vColumna_Inicial) & ")" & vbCrLf & _
                 "- vColumna_LineaSuma = " & vColumna_LineaSuma & _
                 " (Inicial+" & (vColumna_LineaSuma - vColumna_Inicial) & ")" & vbCrLf & vbCrLf & _
                 "Para desactivar este mensaje, cambiar True por False en el c�digo."
    
    MsgBox strMensaje, vbInformation, "Variables de Columnas de Control"
End Sub


Public Function fun815_BorrarColumnasInnecesarias(ByRef ws As Worksheet, _
                                                  ByVal vFila_Inicial As Long, _
                                                  ByVal vFila_Final As Long, _
                                                  ByVal vColumna_Inicial As Long, _
                                                  ByVal vColumna_IdentificadorDeLinea As Long, _
                                                  ByVal vColumna_LineaRepetida As Long, _
                                                  ByVal vColumna_LineaSuma As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun815_BorrarColumnasInnecesarias
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim i As Long
    
    ' Borrar columna identificador de l�nea
    ws.Range(ws.Cells(vFila_Inicial, vColumna_IdentificadorDeLinea), _
             ws.Cells(vFila_Final, vColumna_IdentificadorDeLinea)).Clear
    
    ' Borrar columna l�nea repetida
    ws.Range(ws.Cells(vFila_Inicial, vColumna_LineaRepetida), _
             ws.Cells(vFila_Final, vColumna_LineaRepetida)).Clear
    
    ' Borrar columnas a la izquierda de vColumna_Inicial (excluyendo vColumna_Inicial)
    If vColumna_Inicial > 1 Then
        For i = 1 To vColumna_Inicial - 1
            ws.Range(ws.Cells(vFila_Inicial, i), _
                     ws.Cells(vFila_Final, i)).Clear
        Next i
    End If
    
    ' Borrar columnas a la derecha de vColumna_LineaSuma (excluyendo vColumna_LineaSuma)
    For i = vColumna_LineaSuma + 1 To ws.Columns.Count
        ' Solo limpiar si hay contenido para optimizar rendimiento
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(vFila_Inicial, i), _
                                                         ws.Cells(vFila_Final, i))) > 0 Then
            ws.Range(ws.Cells(vFila_Inicial, i), _
                     ws.Cells(vFila_Final, i)).Clear
        Else
            Exit For ' Si no hay contenido, salir del bucle
        End If
    Next i
    
    fun815_BorrarColumnasInnecesarias = True
    Exit Function
    
GestorErrores:
    fun815_BorrarColumnasInnecesarias = False
End Function


Public Function fun816_FiltrarLineasEspecificas(ByRef ws As Worksheet, _
                                               ByVal vFila_Inicial As Long, _
                                               ByVal vFila_Final As Long, _
                                               ByVal vColumna_Inicial As Long, _
                                               ByVal vColumna_LineaTratada As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun816_FiltrarLineasEspecificas
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim i As Long
    Dim vValor_Columna_Inicial As String
    Dim vValor_Primer_Caracter_Columna_Inicial As String
    Dim vValor_Columna_LineaTratada As String
    Dim blnBorrarLinea As Boolean
    
    ' Recorrer l�neas desde la final hacia la inicial para evitar problemas de �ndices
    For i = vFila_Final To vFila_Inicial Step -1
        
        ' Reinicializar variables para cada l�nea
        vValor_Columna_Inicial = ""
        vValor_Primer_Caracter_Columna_Inicial = ""
        vValor_Columna_LineaTratada = ""
        blnBorrarLinea = False
        
        ' Obtener valor de la primera columna
        vValor_Columna_Inicial = Trim(CStr(ws.Cells(i, vColumna_Inicial).Value))
        
        ' Obtener primer car�cter si hay contenido
        If Len(vValor_Columna_Inicial) > 0 Then
            vValor_Primer_Caracter_Columna_Inicial = Left(vValor_Columna_Inicial, 1)
        Else
            vValor_Primer_Caracter_Columna_Inicial = ""
        End If
        
        ' Obtener valor de columna l�nea tratada
        vValor_Columna_LineaTratada = Trim(CStr(ws.Cells(i, vColumna_LineaTratada).Value))
        
        ' Evaluar criterios para borrar l�nea
        If (vValor_Primer_Caracter_Columna_Inicial = "!") Or _
           (vValor_Columna_Inicial = "") Or _
           (Len(Trim(vValor_Columna_Inicial)) = 0) Or _
           (vValor_Columna_LineaTratada = CONST_TAG_LINEA_TRATADA) Then
            
            blnBorrarLinea = True
        End If
        
        ' Borrar contenido de toda la l�nea si cumple criterios
        If blnBorrarLinea Then
            ws.Rows(i).ClearContents
        End If
        
    Next i
    
    fun816_FiltrarLineasEspecificas = True
    Exit Function
    
GestorErrores:
    fun816_FiltrarLineasEspecificas = False
End Function
