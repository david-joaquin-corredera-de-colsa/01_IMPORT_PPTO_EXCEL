Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_03"

Option Explicit

Public Function fun812_CopiarContenidoCompleto(ByRef wsOrigen As Worksheet, _
                                               ByRef wsDestino As Worksheet) As Boolean
    
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR CORREGIDA: fun812_CopiarContenidoCompleto
    ' Fecha y Hora de Modificación: 2025-06-01 19:34:00 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Copia todo el contenido de una hoja de trabajo a otra hoja de destino
    ' MANTENIENDO LA POSICIÓN ORIGINAL de los datos (ej: si origen está en B2,
    ' destino también estará en B2).
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
    
    ' Calcular celda destino manteniendo posición original
    ' Si el rango origen empieza en B2, el destino también empezará en B2
    strCeldaDestino = wsDestino.Cells(rngUsedOrigen.Row, rngUsedOrigen.Column).Address
    
    ' Copiar manteniendo posición original
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
    ' FUNCIÓN AUXILIAR: fun813_DetectarRangoCompleto
    ' Fecha y Hora de Creación: 2025-06-01 19:20:05 UTC
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
    ' FUNCIÓN AUXILIAR: fun814_MostrarInformacionColumnas
    ' Fecha y Hora de Creación: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    Dim strMensaje As String
    
    strMensaje = "INFORMACIÓN DE VARIABLES DE COLUMNAS DE CONTROL" & vbCrLf & vbCrLf & _
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
                 "Para desactivar este mensaje, cambiar True por False en el código."
    
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
    ' FUNCIÓN AUXILIAR: fun815_BorrarColumnasInnecesarias
    ' Fecha y Hora de Creación: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim i As Long
    
    ' Borrar columna identificador de línea
    ws.Range(ws.Cells(vFila_Inicial, vColumna_IdentificadorDeLinea), _
             ws.Cells(vFila_Final, vColumna_IdentificadorDeLinea)).Clear
    
    ' Borrar columna línea repetida
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
    ' FUNCIÓN AUXILIAR: fun816_FiltrarLineasEspecificas
    ' Fecha y Hora de Creación: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim i As Long
    Dim vValor_Columna_Inicial As String
    Dim vValor_Primer_Caracter_Columna_Inicial As String
    Dim vValor_Columna_LineaTratada As String
    Dim blnBorrarLinea As Boolean
    
    ' Recorrer líneas desde la final hacia la inicial para evitar problemas de índices
    For i = vFila_Final To vFila_Inicial Step -1
        
        ' Reinicializar variables para cada línea
        vValor_Columna_Inicial = ""
        vValor_Primer_Caracter_Columna_Inicial = ""
        vValor_Columna_LineaTratada = ""
        blnBorrarLinea = False
        
        ' Obtener valor de la primera columna
        vValor_Columna_Inicial = Trim(CStr(ws.Cells(i, vColumna_Inicial).Value))
        
        ' Obtener primer carácter si hay contenido
        If Len(vValor_Columna_Inicial) > 0 Then
            vValor_Primer_Caracter_Columna_Inicial = Left(vValor_Columna_Inicial, 1)
        Else
            vValor_Primer_Caracter_Columna_Inicial = ""
        End If
        
        ' Obtener valor de columna línea tratada
        vValor_Columna_LineaTratada = Trim(CStr(ws.Cells(i, vColumna_LineaTratada).Value))
        
        ' Evaluar criterios para borrar línea
        If (vValor_Primer_Caracter_Columna_Inicial = "!") Or _
           (vValor_Columna_Inicial = "") Or _
           (Len(Trim(vValor_Columna_Inicial)) = 0) Or _
           (vValor_Columna_LineaTratada = CONST_TAG_LINEA_TRATADA) Then
            
            blnBorrarLinea = True
        End If
        
        ' Borrar contenido de toda la línea si cumple criterios
        If blnBorrarLinea Then
            ws.Rows(i).ClearContents
        End If
        
    Next i
    
    fun816_FiltrarLineasEspecificas = True
    Exit Function
    
GestorErrores:
    fun816_FiltrarLineasEspecificas = False
End Function

Public Function fun817_CopiarContenidoCompleto(ByRef wsOrigen As Worksheet, _
                                               ByRef wsDestino As Worksheet) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun817_CopiarContenidoCompleto
    ' Fecha y Hora de Creación: 2025-06-01 21:52:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Copia todo el contenido de una hoja de trabajo a otra hoja de destino
    ' MANTENIENDO LA POSICIÓN ORIGINAL de los datos (ej: si origen está en B2,
    ' destino también estará en B2).
    '
    ' Parámetros:
    ' - wsOrigen: Hoja de trabajo origen
    ' - wsDestino: Hoja de trabajo destino
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para el procesamiento
    Dim rngUsedOrigen As Range
    Dim strCeldaDestino As String
    
    ' Inicialización
    strFuncion = "fun817_CopiarContenidoCompleto" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun817_CopiarContenidoCompleto"
    fun817_CopiarContenidoCompleto = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If wsOrigen Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 201, strFuncion, _
            "Hoja de origen no válida"
    End If
    
    If wsDestino Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 202, strFuncion, _
            "Hoja de destino no válida"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Limpiar hoja destino
    '--------------------------------------------------------------------------
    lngLineaError = 40
    If Not fun801_LimpiarHoja(wsDestino.Name) Then
        Err.Raise ERROR_BASE_IMPORT + 203, strFuncion, _
            "Error al limpiar hoja de destino: " & wsDestino.Name
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Verificar que hay contenido en la hoja origen
    '--------------------------------------------------------------------------
    lngLineaError = 50
    If wsOrigen.UsedRange Is Nothing Then
        ' No hay contenido, pero no es error
        fun817_CopiarContenidoCompleto = True
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Obtener rango usado de origen y calcular destino
    '--------------------------------------------------------------------------
    lngLineaError = 60
    Set rngUsedOrigen = wsOrigen.UsedRange
    
    ' Calcular celda destino manteniendo posición original
    ' Si el rango origen empieza en B2, el destino también empezará en B2
    strCeldaDestino = wsDestino.Cells(rngUsedOrigen.Row, rngUsedOrigen.Column).Address
    
    '--------------------------------------------------------------------------
    ' 5. Realizar la copia manteniendo posición original
    '--------------------------------------------------------------------------
    lngLineaError = 70
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Copiar contenido
    rngUsedOrigen.Copy wsDestino.Range(strCeldaDestino)
    
    ' Limpiar portapapeles para liberar memoria
    Application.CutCopyMode = False
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 6. Finalización exitosa
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun817_CopiarContenidoCompleto = True
    Exit Function

GestorErrores:
    ' Restaurar configuración
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun817_CopiarContenidoCompleto = False
End Function

Public Function fun818_BorrarColumnaLineaSuma(ByRef ws As Worksheet, _
                                             ByVal vColumna_LineaSuma As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun818_BorrarColumnaLineaSuma
    ' Fecha y Hora de Creación: 2025-06-02 03:27:31 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Borra todo el contenido y formatos de la columna vColumna_LineaSuma
    ' en toda la hoja de trabajo especificada.
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde borrar la columna
    ' - vColumna_LineaSuma: Número de columna a borrar
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Inicialización
    strFuncion = "fun818_BorrarColumnaLineaSuma" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun818_BorrarColumnaLineaSuma"
    fun818_BorrarColumnaLineaSuma = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 301, strFuncion, _
            "Hoja de trabajo no válida"
    End If
    
    If vColumna_LineaSuma < 1 Or vColumna_LineaSuma > 16384 Then
        Err.Raise ERROR_BASE_IMPORT + 302, strFuncion, _
            "Número de columna fuera de rango: " & vColumna_LineaSuma
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Borrar contenido y formatos de toda la columna
    '--------------------------------------------------------------------------
    lngLineaError = 40
    With ws.Columns(vColumna_LineaSuma)
        .Clear
    End With
    
    fun818_BorrarColumnaLineaSuma = True
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun818_BorrarColumnaLineaSuma = False
End Function

Public Function fun819_DetectarPrimeraFilaContenido(ByRef ws As Worksheet, _
                                                   ByVal vColumna_Inicial As Long, _
                                                   ByRef vFila_Inicial_HojaLimpia As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun819_DetectarPrimeraFilaContenido
    ' Fecha y Hora de Creación: 2025-06-02 03:27:31 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Detecta la primera fila que contiene datos en la columna inicial especificada
    ' después de las operaciones de limpieza.
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde detectar
    ' - vColumna_Inicial: Columna donde buscar contenido
    ' - vFila_Inicial_HojaLimpia: Variable donde almacenar el resultado
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para búsqueda
    Dim rngBusqueda As Range
    Dim rngEncontrado As Range
    
    ' Inicialización
    strFuncion = "fun819_DetectarPrimeraFilaContenido" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun819_DetectarPrimeraFilaContenido"
    fun819_DetectarPrimeraFilaContenido = False
    lngLineaError = 0
    vFila_Inicial_HojaLimpia = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 401, strFuncion, _
            "Hoja de trabajo no válida"
    End If
    
    If vColumna_Inicial < 1 Or vColumna_Inicial > 16384 Then
        Err.Raise ERROR_BASE_IMPORT + 402, strFuncion, _
            "Número de columna fuera de rango: " & vColumna_Inicial
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Buscar primera celda con contenido en la columna especificada
    '--------------------------------------------------------------------------
    lngLineaError = 40
    Set rngBusqueda = ws.Columns(vColumna_Inicial)
    
    ' Buscar primera celda con contenido
    Set rngEncontrado = rngBusqueda.Find(What:="*", _
                                        After:=rngBusqueda.Cells(rngBusqueda.Cells.Count), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByRows, _
                                        SearchDirection:=xlNext)
    
    '--------------------------------------------------------------------------
    ' 3. Procesar resultado
    '--------------------------------------------------------------------------
    lngLineaError = 50
    If Not rngEncontrado Is Nothing Then
        vFila_Inicial_HojaLimpia = rngEncontrado.Row
        fun819_DetectarPrimeraFilaContenido = True
    Else
        ' No se encontró contenido, asignar fila por defecto
        vFila_Inicial_HojaLimpia = 3 ' Fila 3 por defecto para dejar espacio para headers
        fun819_DetectarPrimeraFilaContenido = True
    End If
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    vFila_Inicial_HojaLimpia = 0
    fun819_DetectarPrimeraFilaContenido = False
End Function

Public Function fun820_AnadirHeadersIdentificativos(ByRef ws As Worksheet, _
                                                   ByVal vFila_Inicial_HojaLimpia As Long, _
                                                   ByVal vColumna_Inicial As Long, _
                                                   ByRef vScenario_HEnvio As String, _
                                                   ByRef vYear_HEnvio As String, _
                                                   ByRef vEntity_HEnvio As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun820_AnadirHeadersIdentificativos
    ' Fecha y Hora de Creación: 2025-06-02 03:27:31 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Añade headers identificativos en la fila vFila_Inicial_HojaLimpia-1
    ' con los valores especificados para las columnas 0 a 10.
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde añadir headers
    ' - vFila_Inicial_HojaLimpia: Fila de referencia para calcular posición
    ' - vColumna_Inicial: Columna inicial donde comenzar
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para posicionamiento
    Dim lngFilaHeader As Long
    
    ' Inicialización
    strFuncion = "fun820_AnadirHeadersIdentificativos" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun820_AnadirHeadersIdentificativos"
    fun820_AnadirHeadersIdentificativos = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 501, strFuncion, _
            "Hoja de trabajo no válida"
    End If
    
    If vFila_Inicial_HojaLimpia < 2 Then
        Err.Raise ERROR_BASE_IMPORT + 502, strFuncion, _
            "Fila inicial debe ser mayor a 1 para poder añadir headers"
    End If
    
    If vColumna_Inicial < 1 Or vColumna_Inicial > 16384 Then
        Err.Raise ERROR_BASE_IMPORT + 503, strFuncion, _
            "Número de columna fuera de rango: " & vColumna_Inicial
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Calcular fila donde añadir headers identificativos
    '--------------------------------------------------------------------------
    lngLineaError = 40
    lngFilaHeader = vFila_Inicial_HojaLimpia - 1
    
    '--------------------------------------------------------------------------
    ' 3. Añadir headers identificativos (columnas 0 a 10)
    '--------------------------------------------------------------------------
    lngLineaError = 50
    With ws
        .Cells(lngFilaHeader, vColumna_Inicial + 0).Value = "Budget_OS"
        .Cells(lngFilaHeader, vColumna_Inicial + 1).Value = "2031"
        .Cells(lngFilaHeader, vColumna_Inicial + 2).Value = "YTD"
        .Cells(lngFilaHeader, vColumna_Inicial + 3).Value = "GR_HOLD"
        .Cells(lngFilaHeader, vColumna_Inicial + 4).Value = "<Entity Currency>"
        .Cells(lngFilaHeader, vColumna_Inicial + 5).Value = "RESULT"
        .Cells(lngFilaHeader, vColumna_Inicial + 6).Value = "[ICP Top]"
        .Cells(lngFilaHeader, vColumna_Inicial + 7).Value = "TotC1"
        .Cells(lngFilaHeader, vColumna_Inicial + 8).Value = "TotC2"
        .Cells(lngFilaHeader, vColumna_Inicial + 9).Value = "TotC3"
        .Cells(lngFilaHeader, vColumna_Inicial + 10).Value = "TotC4"
    End With
    
    'En vez de a Entity "GR_HOLD" vamos a poner en la primera línea la misma Entity que en el resto del fichero
    '   asi que tomamos la Entity de la linea siguiente a la del "Header" (esto es lngFilaHeader + 1)
    '   Es decir, tomamos la Entity de la primera linea de datos importada
    vEntity_HEnvio = ws.Cells(lngFilaHeader + 1, vColumna_Inicial + 3).Value
    ws.Cells(lngFilaHeader, vColumna_Inicial + 3).Value = vEntity_HEnvio 'ws.Cells(lngFilaHeader + 1, vColumna_Inicial + 3).Value
    'Y para el año y el escenario hacemos algo parecido
    '   El Escenario lo tomamos de la constante CONST_ESCENARIO_ADMITIDO
    vScenario_HEnvio = CONST_ESCENARIO_ADMITIDO
    ws.Cells(lngFilaHeader, vColumna_Inicial + 0).Value = vScenario_HEnvio
    '   El Año lo tomamos de la primera linea de datos importada (esto es lngFilaHeader + 1)
    vYear_HEnvio = ws.Cells(lngFilaHeader + 1, vColumna_Inicial + 1).Value
    ws.Cells(lngFilaHeader, vColumna_Inicial + 1).Value = vYear_HEnvio
    
    fun820_AnadirHeadersIdentificativos = True
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun820_AnadirHeadersIdentificativos = False
End Function

Public Function fun821_AnadirHeadersMeses(ByRef ws As Worksheet, _
                                         ByVal vFila_Inicial_HojaLimpia As Long, _
                                         ByVal vColumna_Inicial As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun821_AnadirHeadersMeses
    ' Fecha y Hora de Creación: 2025-06-02 03:27:31 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Añade headers de meses en la fila vFila_Inicial_HojaLimpia-2
    ' con los valores M01 a M12 para las columnas 11 a 22.
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde añadir headers
    ' - vFila_Inicial_HojaLimpia: Fila de referencia para calcular posición
    ' - vColumna_Inicial: Columna inicial donde comenzar
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para posicionamiento
    Dim lngFilaHeader As Long
    
    ' Inicialización
    strFuncion = "fun821_AnadirHeadersMeses" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun821_AnadirHeadersMeses"
    fun821_AnadirHeadersMeses = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 601, strFuncion, _
            "Hoja de trabajo no válida"
    End If
    
    If vFila_Inicial_HojaLimpia < 3 Then
        Err.Raise ERROR_BASE_IMPORT + 602, strFuncion, _
            "Fila inicial debe ser mayor a 2 para poder añadir headers de meses"
    End If
    
    If vColumna_Inicial < 1 Or vColumna_Inicial > 16384 Then
        Err.Raise ERROR_BASE_IMPORT + 603, strFuncion, _
            "Número de columna fuera de rango: " & vColumna_Inicial
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Calcular fila donde añadir headers de meses
    '--------------------------------------------------------------------------
    lngLineaError = 40
    lngFilaHeader = vFila_Inicial_HojaLimpia - 2
    
    '--------------------------------------------------------------------------
    ' 3. Añadir headers de meses (columnas 11 a 22)
    '--------------------------------------------------------------------------
    lngLineaError = 50
    With ws
        .Cells(lngFilaHeader, vColumna_Inicial + 11).Value = "M01"
        .Cells(lngFilaHeader, vColumna_Inicial + 12).Value = "M02"
        .Cells(lngFilaHeader, vColumna_Inicial + 13).Value = "M03"
        .Cells(lngFilaHeader, vColumna_Inicial + 14).Value = "M04"
        .Cells(lngFilaHeader, vColumna_Inicial + 15).Value = "M05"
        .Cells(lngFilaHeader, vColumna_Inicial + 16).Value = "M06"
        .Cells(lngFilaHeader, vColumna_Inicial + 17).Value = "M07"
        .Cells(lngFilaHeader, vColumna_Inicial + 18).Value = "M08"
        .Cells(lngFilaHeader, vColumna_Inicial + 19).Value = "M09"
        .Cells(lngFilaHeader, vColumna_Inicial + 20).Value = "M10"
        .Cells(lngFilaHeader, vColumna_Inicial + 21).Value = "M11"
        .Cells(lngFilaHeader, vColumna_Inicial + 22).Value = "M12"
    End With
    
    fun821_AnadirHeadersMeses = True
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun821_AnadirHeadersMeses = False
End Function



Public Function fun822_DetectarRangoCompletoHoja(ByRef ws As Worksheet, _
                                                ByRef vFila_Inicial As Long, _
                                                ByRef vFila_Final As Long, _
                                                ByRef vColumna_Inicial As Long, _
                                                 ByRef vColumna_Final As Long) As Boolean
    
    'Posible redundancia con funcion fun813_DetectarRangoCompleto
    '******************************************************************************
    ' FUNCIÓN AUXILIAR MEJORADA: fun822_DetectarRangoCompletoHoja
    ' Fecha y Hora de Creación: 2025-06-03 03:19:45 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Detecta el rango completo de datos en una hoja de trabajo específica
    ' basándose en palabras clave definidas en variables globales.
    ' Reutilizada por F007_Copiar_Datos_de_Comprobacion_a_Envio
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo a analizar
    ' - vFila_Inicial: Variable donde almacenar primera fila con palabra clave
    ' - vFila_Final: Variable donde almacenar última fila con palabra clave
    ' - vColumna_Inicial: Variable donde almacenar primera columna con palabra clave
    ' - vColumna_Final: Variable donde almacenar última columna con palabra clave
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Variables para búsqueda
    Dim rngBusqueda As Range
    Dim rngEncontrado As Range
    Dim strPalabraBuscar As String
    
    ' Inicialización
    strFuncion = "fun822_DetectarRangoCompletoHoja" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun822_DetectarRangoCompletoHoja"
    lngLineaError = 0
    
    ' Inicializar valores por defecto
    vFila_Inicial = 0
    vFila_Final = 0
    vColumna_Inicial = 0
    vColumna_Final = 0
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetro de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If ws Is Nothing Then
        fun822_DetectarRangoCompletoHoja = False
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Buscar PRIMERA FILA con palabra clave
    '--------------------------------------------------------------------------
    lngLineaError = 40
    strPalabraBuscar = UCase(Trim(vPalabraClave_PrimeraFila))
    
    If Len(strPalabraBuscar) > 0 Then
        Set rngBusqueda = ws.UsedRange
        If Not rngBusqueda Is Nothing Then
            Set rngEncontrado = rngBusqueda.Find(What:=strPalabraBuscar, _
                                                LookIn:=xlValues, _
                                                LookAt:=xlWhole, _
                                                SearchOrder:=xlByRows, _
                                                SearchDirection:=xlNext, _
                                                MatchCase:=False)
            If Not rngEncontrado Is Nothing Then
                vFila_Inicial = rngEncontrado.Row
            End If
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Buscar PRIMERA COLUMNA con palabra clave
    '--------------------------------------------------------------------------
    lngLineaError = 50
    strPalabraBuscar = UCase(Trim(vPalabraClave_PrimeraColumna))
    
    If Len(strPalabraBuscar) > 0 Then
        Set rngBusqueda = ws.UsedRange
        If Not rngBusqueda Is Nothing Then
            Set rngEncontrado = rngBusqueda.Find(What:=strPalabraBuscar, _
                                                LookIn:=xlValues, _
                                                LookAt:=xlWhole, _
                                                SearchOrder:=xlByColumns, _
                                                SearchDirection:=xlNext, _
                                                MatchCase:=False)
            If Not rngEncontrado Is Nothing Then
                vColumna_Inicial = rngEncontrado.Column
            End If
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Buscar ÚLTIMA FILA con palabra clave
    '--------------------------------------------------------------------------
    lngLineaError = 60
    strPalabraBuscar = UCase(Trim(vPalabraClave_UltimaFila))
    
    If Len(strPalabraBuscar) > 0 Then
        Set rngBusqueda = ws.UsedRange
        If Not rngBusqueda Is Nothing Then
            Set rngEncontrado = rngBusqueda.Find(What:=strPalabraBuscar, _
                                                LookIn:=xlValues, _
                                                LookAt:=xlWhole, _
                                                SearchOrder:=xlByRows, _
                                                SearchDirection:=xlPrevious, _
                                                MatchCase:=False)
            If Not rngEncontrado Is Nothing Then
                vFila_Final = rngEncontrado.Row
            End If
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Buscar ÚLTIMA COLUMNA con palabra clave
    '--------------------------------------------------------------------------
    lngLineaError = 70
    strPalabraBuscar = UCase(Trim(vPalabraClave_UltimaColumna))
    
    If Len(strPalabraBuscar) > 0 Then
        Set rngBusqueda = ws.UsedRange
        If Not rngBusqueda Is Nothing Then
            Set rngEncontrado = rngBusqueda.Find(What:=strPalabraBuscar, _
                                                LookIn:=xlValues, _
                                                LookAt:=xlWhole, _
                                                SearchOrder:=xlByColumns, _
                                                SearchDirection:=xlPrevious, _
                                                MatchCase:=False)
            If Not rngEncontrado Is Nothing Then
                vColumna_Final = rngEncontrado.Column
            End If
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Validar que se encontraron todos los rangos
    '--------------------------------------------------------------------------
    lngLineaError = 80
    If vFila_Inicial > 0 And vFila_Final > 0 And vColumna_Inicial > 0 And vColumna_Final > 0 Then
        ' Validar lógica de rangos
        If vFila_Inicial <= vFila_Final And vColumna_Inicial <= vColumna_Final Then
            fun822_DetectarRangoCompletoHoja = True
        Else
            ' Los rangos no son lógicos, intentar corregir
            If vFila_Inicial > vFila_Final Then
                Dim tempFila As Long
                tempFila = vFila_Inicial
                vFila_Inicial = vFila_Final
                vFila_Final = tempFila
            End If
            
            If vColumna_Inicial > vColumna_Final Then
                Dim tempColumna As Long
                tempColumna = vColumna_Inicial
                vColumna_Inicial = vColumna_Final
                vColumna_Final = tempColumna
            End If
            
            fun822_DetectarRangoCompletoHoja = True
        End If
    Else
        fun822_DetectarRangoCompletoHoja = False
    End If
    
    Exit Function
    
GestorErrores:

    ' Construir mensaje de error detallado
    Dim strMensajeError As String
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja: " & ws.Name
    
    fun801_LogMessage strMensajeError, True
    
    vFila_Inicial = 0
    vFila_Final = 0
    vColumna_Inicial = 0
    vColumna_Final = 0
    fun822_DetectarRangoCompletoHoja = False
    
End Function

Public Function fun823_CopiarSoloValores(ByRef rngOrigen As Range, _
                                        ByRef rngDestino As Range) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_CopiarSoloValores
    ' Fecha y Hora de Creación: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Copia únicamente los valores (sin formatos) de un rango origen a un rango destino
    ' Compatible con repositorios OneDrive, SharePoint y Teams
    '
    ' Parámetros:
    ' - rngOrigen: Rango de celdas origen
    ' - rngDestino: Rango de celdas destino
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Validar parámetros
    If rngOrigen Is Nothing Or rngDestino Is Nothing Then
        fun823_CopiarSoloValores = False
        Exit Function
    End If
    
    ' Configurar entorno para optimizar rendimiento
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Copiar y pegar solo valores (método compatible Excel 97-365)
    rngOrigen.Copy
    rngDestino.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Restaurar configuración
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    fun823_CopiarSoloValores = True
    Exit Function
    
GestorErrores:
    ' Limpiar portapapeles y restaurar configuración
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    fun823_CopiarSoloValores = False
End Function
