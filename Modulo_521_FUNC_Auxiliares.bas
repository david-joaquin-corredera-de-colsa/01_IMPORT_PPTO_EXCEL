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
    ' FUNCI�N: fun802_SeleccionarArchivo
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Proporciona una interfaz de usuario intuitiva para seleccionar archivos de
    ' texto (TXT y CSV) mediante un cuadro de di�logo est�ndar de Windows. Facilita
    ' la selecci�n de archivos de datos para importaci�n en el sistema de presupuestos.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = strPrompt
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt;*.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            fun802_SeleccionarArchivo = .SelectedItems(1)
        Else
            fun802_SeleccionarArchivo = ""
        End If
    End With
    Exit Function
    
GestorErrores:
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


Public Function fun805_ConvertirTextoAColumnas(ByRef rngConversion As Range, _
                                             ByVal strDelimitador As String) As Boolean
    
    '*******************************************************************************
    ' Funci�n: fun805_ConvertirTextoAColumnas
    '
    ' Descripci�n:
    ' Convierte texto en columnas utilizando un delimitador espec�fico. Es similar
    ' a la funcionalidad "Texto en columnas" de Excel, pero automatizada mediante VBA
    ' y con configuraciones predefinidas para los tipos de datos de las columnas resultantes.
    '
    ' Par�metros:
    ' - rngConversion (Range): Rango de celdas que contiene el texto a convertir
    ' - strDelimitador (String): Car�cter que se utilizar� como delimitador
    '
    ' Retorno:
    ' - Boolean: True si la conversi�n se realiz� correctamente, False en caso de error
    '
    ' Pasos que realiza:
    ' 1. Desactiva la actualizaci�n de pantalla y eventos para mejorar rendimiento
    ' 2. Crea un array para definir los tipos de datos de cada columna resultante
    ' 3. Configura las columnas 1-11 como tipo Texto (c�digo 2)
    ' 4. Configura las columnas 12-23 como tipo General (c�digo 1)
    ' 5. Ejecuta la conversi�n usando el m�todo TextToColumns con el delimitador especificado
    ' 6. Restaura la actualizaci�n de pantalla y eventos
    ' 7. Devuelve True si todo fue exitoso o False si hubo alg�n error
    '
    ' Configuraci�n de columnas:
    ' - Columnas 1-11: Configuradas como Texto (evita conversiones autom�ticas)
    ' - Columnas 12-23: Configuradas como General (permite conversiones autom�ticas)
    '
    ' Ejemplo de uso:
    ' Dim resultado As Boolean
    ' resultado = fun805_ConvertirTextoAColumnas(Range("A1:A10"), ";")
    '
    ' Notas:
    ' - La funci�n desactiva temporalmente la actualizaci�n de pantalla y eventos
    ' - Incluye manejo de errores para garantizar que se restaure el entorno
    '
    ' Fecha: 2025-05-28 18:06:56
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Versi�n: 1.0
    '*******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Configurar entorno
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Array para tipos de columna (1=General, 2=Texto)
    Dim varTipos As Variant
    Dim i As Long
    
    ReDim varTipos(1 To 23, 1 To 2)
    
    ' Configurar tipos de columna
    For i = 1 To 11    ' Columnas 1-11: Texto
        varTipos(i, 1) = i
        varTipos(i, 2) = 2
    Next i
    
    For i = 12 To 23   ' Columnas 12-23: General
        varTipos(i, 1) = i
        varTipos(i, 2) = 1
    Next i
    
    ' Ejecutar conversi�n
    rngConversion.TextToColumns _
        Destination:=rngConversion, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=True, _
        OtherChar:=strDelimitador, _
        FieldInfo:=varTipos
    
    ' Restaurar entorno
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    fun805_ConvertirTextoAColumnas = True
    Exit Function
    
GestorErrores:
    ' Restaurar entorno
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    fun805_ConvertirTextoAColumnas = False
End Function

