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
    strFuncion = "fun817_CopiarContenidoCompleto"
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
    strFuncion = "fun818_BorrarColumnaLineaSuma"
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
    strFuncion = "fun819_DetectarPrimeraFilaContenido"
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
                                                   ByVal vColumna_Inicial As Long) As Boolean
    
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
    strFuncion = "fun820_AnadirHeadersIdentificativos"
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
    strFuncion = "fun821_AnadirHeadersMeses"
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
    strFuncion = "fun822_DetectarRangoCompletoHoja"
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

Public Function fun824_LimpiarFilasExcedentes(ByRef ws As Worksheet, _
                                             ByVal vFila_Final_Limite As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun824_LimpiarFilasExcedentes
    ' Fecha y Hora de Creación: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Limpia todas las filas que estén por encima del límite especificado
    ' Borra tanto contenido como formatos para optimizar el archivo
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde limpiar
    ' - vFila_Final_Limite: Número de fila límite (se borran filas superiores)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngUltimaFilaConDatos As Long
    
    ' Validar parámetros
    If ws Is Nothing Then
        fun824_LimpiarFilasExcedentes = False
        Exit Function
    End If
    
    If vFila_Final_Limite < 1 Then
        fun824_LimpiarFilasExcedentes = False
        Exit Function
    End If
    
    ' Obtener última fila con datos (método compatible Excel 97-365)
    lngUltimaFilaConDatos = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Si hay filas excedentes, limpiarlas completamente
    If lngUltimaFilaConDatos > vFila_Final_Limite Then
        Application.ScreenUpdating = False
        
        ' Limpiar contenido y formatos (compatible Excel 97-365)
        ws.Range(ws.Cells(vFila_Final_Limite + 1, 1), _
                 ws.Cells(lngUltimaFilaConDatos, ws.Columns.Count)).Clear
        
        Application.ScreenUpdating = True
    End If
    
    fun824_LimpiarFilasExcedentes = True
    Exit Function
    
GestorErrores:
    Application.ScreenUpdating = True
    fun824_LimpiarFilasExcedentes = False
End Function

Public Function fun825_LimpiarColumnasExcedentes(ByRef ws As Worksheet, _
                                                ByVal vColumna_Final_Limite As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun825_LimpiarColumnasExcedentes
    ' Fecha y Hora de Creación: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Limpia todas las columnas que estén por encima del límite especificado
    ' Borra tanto contenido como formatos para optimizar el archivo
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde limpiar
    ' - vColumna_Final_Limite: Número de columna límite (se borran columnas superiores)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngUltimaColumnaConDatos As Long
    
    ' Validar parámetros
    If ws Is Nothing Then
        fun825_LimpiarColumnasExcedentes = False
        Exit Function
    End If
    
    If vColumna_Final_Limite < 1 Then
        fun825_LimpiarColumnasExcedentes = False
        Exit Function
    End If
    
    ' Obtener última columna con datos (método compatible Excel 97-365)
    lngUltimaColumnaConDatos = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Si hay columnas excedentes, limpiarlas completamente
    If lngUltimaColumnaConDatos > vColumna_Final_Limite Then
        Application.ScreenUpdating = False
        
        ' Limpiar contenido y formatos (compatible Excel 97-365)
        ws.Range(ws.Cells(1, vColumna_Final_Limite + 1), _
                 ws.Cells(ws.Rows.Count, lngUltimaColumnaConDatos)).Clear
        
        Application.ScreenUpdating = True
    End If
    
    fun825_LimpiarColumnasExcedentes = True
    Exit Function
    
GestorErrores:
    Application.ScreenUpdating = True
    fun825_LimpiarColumnasExcedentes = False
End Function

Public Function fun826_ConfigurarPalabrasClave(Optional ByVal strPrimeraFila As String = "", _
                                              Optional ByVal strPrimeraColumna As String = "", _
                                              Optional ByVal strUltimaFila As String = "", _
                                              Optional ByVal strUltimaColumna As String = "") As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun826_ConfigurarPalabrasClave
    ' Fecha y Hora de Creación: 2025-06-03 03:19:45 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Permite configurar las palabras clave utilizadas para detectar rangos
    ' de datos en las hojas de trabajo.
    '
    ' Parámetros (todos opcionales):
    ' - strPrimeraFila: Palabra clave para buscar primera fila
    ' - strPrimeraColumna: Palabra clave para buscar primera columna
    ' - strUltimaFila: Palabra clave para buscar última fila
    ' - strUltimaColumna: Palabra clave para buscar última columna
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Solo actualizar las variables que se proporcionen
    If Len(Trim(strPrimeraFila)) > 0 Then
        vPalabraClave_PrimeraFila = Trim(strPrimeraFila)
    End If
    
    If Len(Trim(strPrimeraColumna)) > 0 Then
        vPalabraClave_PrimeraColumna = Trim(strPrimeraColumna)
    End If
    
    If Len(Trim(strUltimaFila)) > 0 Then
        vPalabraClave_UltimaFila = Trim(strUltimaFila)
    End If
    
    If Len(Trim(strUltimaColumna)) > 0 Then
        vPalabraClave_UltimaColumna = Trim(strUltimaColumna)
    End If
    
    fun826_ConfigurarPalabrasClave = True
    Exit Function
    
GestorErrores:
    fun826_ConfigurarPalabrasClave = False
End Function

Public Function fun821_EsHojaSistema(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun821_EsHojaSistema
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Verifica si una hoja es del sistema (no debe modificarse)
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Select Case strNombreHoja
        Case "00_Ejecutar_Procesos", "01_Inventario", "02_Log", _
             "05_Username", "06_Delimitadores_Originales"
            fun821_EsHojaSistema = True
        Case Else
            fun821_EsHojaSistema = False
    End Select
    
    Exit Function
    
GestorErrores:
    fun821_EsHojaSistema = False
End Function

Public Function fun822_EsHojaImportacion(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun822_EsHojaImportacion
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Verifica si una hoja es de importación, working o comprobación
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    If Left(strNombreHoja, 7) = "Import_" Or _
       Left(strNombreHoja, 15) = "Import_Working_" Or _
       Left(strNombreHoja, 15) = "Del_Prev_Envio_" Or _
       Left(strNombreHoja, 15) = "Import_Comprob_" Then
        fun822_EsHojaImportacion = True
    Else
        fun822_EsHojaImportacion = False
    End If
    
    Exit Function
    
GestorErrores:
    fun822_EsHojaImportacion = False
End Function

Public Function fun823_OcultarHojaSiVisible(ByRef ws As Worksheet) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_OcultarHojaSiVisible
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Oculta una hoja si está visible
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    If ws.Visible = xlSheetVisible Then
        ws.Visible = xlSheetHidden
        fun823_OcultarHojaSiVisible = True
    Else
        fun823_OcultarHojaSiVisible = False
    End If
    
    Exit Function
    
GestorErrores:
    fun823_OcultarHojaSiVisible = False
End Function

Public Function fun824_EsHojaEnvio(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun824_EsHojaEnvio
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Verifica si una hoja es de envío
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    If Left(strNombreHoja, 13) = "Import_Envio_" Then
        fun824_EsHojaEnvio = True
    Else
        fun824_EsHojaEnvio = False
    End If
    
    Exit Function
    
GestorErrores:
    fun824_EsHojaEnvio = False
End Function

Public Function fun825_ProcesarHojasEnvio(ByRef intHojasOcultadas As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun825_ProcesarHojasEnvio
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Procesa hojas de envío, manteniendo visibles las 3 más recientes
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim arrHojasEnvio() As String
    Dim arrFechas() As String
    Dim i As Long, j As Long
    Dim intContador As Integer
    Dim strTemp As String
    Dim strTempFecha As String
    
    ' Recopilar hojas de envío
    intContador = 0
    For i = 1 To ThisWorkbook.Worksheets.Count
        If fun824_EsHojaEnvio(ThisWorkbook.Worksheets(i).Name) Then
            intContador = intContador + 1
            ReDim Preserve arrHojasEnvio(1 To intContador)
            ReDim Preserve arrFechas(1 To intContador)
            
            arrHojasEnvio(intContador) = ThisWorkbook.Worksheets(i).Name
            arrFechas(intContador) = fun826_ExtraerFechaHoja(ThisWorkbook.Worksheets(i).Name)
        End If
    Next i
    
    ' Ordenar por fecha (burbuja simple para compatibilidad)
    For i = 1 To intContador - 1
        For j = i + 1 To intContador
            If arrFechas(i) < arrFechas(j) Then
                ' Intercambiar fechas
                strTempFecha = arrFechas(i)
                arrFechas(i) = arrFechas(j)
                arrFechas(j) = strTempFecha
                
                ' Intercambiar nombres
                strTemp = arrHojasEnvio(i)
                arrHojasEnvio(i) = arrHojasEnvio(j)
                arrHojasEnvio(j) = strTemp
            End If
        Next j
    Next i
    
    ' Ocultar hojas excepto las 3 primeras (más recientes)
    intHojasOcultadas = 0
    For i = 4 To intContador
        If fun823_OcultarHojaSiVisible(ThisWorkbook.Worksheets(arrHojasEnvio(i))) Then
            intHojasOcultadas = intHojasOcultadas + 1
        End If
    Next i
    
    fun825_ProcesarHojasEnvio = True
    Exit Function
    
GestorErrores:
    fun825_ProcesarHojasEnvio = False
End Function

Public Function fun826_ExtraerFechaHoja(ByVal strNombreHoja As String) As String
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun826_ExtraerFechaHoja
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Extrae la fecha/hora del nombre de hoja formato _yyyyMMdd_hhmmss
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim intPos As Integer
    Dim strSufijo As String
    
    ' Buscar la última aparición de "_" para encontrar el sufijo de fecha
    intPos = InStrRev(strNombreHoja, "_")
    
    If intPos > 0 And intPos < Len(strNombreHoja) Then
        strSufijo = Mid(strNombreHoja, intPos + 1)
        ' Verificar que tiene formato de fecha/hora (15 caracteres: yyyyMMdd_hhmmss)
        If Len(strSufijo) = 15 And Mid(strSufijo, 9, 1) = "_" Then
            fun826_ExtraerFechaHoja = strSufijo
        Else
            fun826_ExtraerFechaHoja = "00000000_000000"
        End If
    Else
        fun826_ExtraerFechaHoja = "00000000_000000"
    End If
    
    Exit Function
    
GestorErrores:
    fun826_ExtraerFechaHoja = "00000000_000000"
End Function

Public Function fun822_EsHojaImportacionSinEnvio(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR CORREGIDA: fun822_EsHojaImportacionSinEnvio
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Fecha y Hora de Modificación: 2025-06-03 04:36:36 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Verifica si una hoja es de importación, working o comprobación
    ' PERO NO de envío (las hojas de envío se procesan por separado)
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Verificar hojas Import_ pero NO Import_Envio_
    If Left(strNombreHoja, 7) = "Import_" Then
        ' Si es Import_Envio_, retornar False (no procesar aquí)
        If Left(strNombreHoja, 13) = "Import_Envio_" Then
            fun822_EsHojaImportacionSinEnvio = False
        ' Si es Import_Working_ o Import_Comprob_, retornar True
        ElseIf Left(strNombreHoja, 15) = "Import_Working_" Or _
               Left(strNombreHoja, 15) = "Del_Prev_Envio_" Or _
               Left(strNombreHoja, 15) = "Import_Comprob_" Then
            fun822_EsHojaImportacionSinEnvio = True
        ' Si es solo Import_ (sin sufijos específicos), retornar True
        Else
            fun822_EsHojaImportacionSinEnvio = True
        End If
    Else
        fun822_EsHojaImportacionSinEnvio = False
    End If
    
    Exit Function
    
GestorErrores:
    fun822_EsHojaImportacionSinEnvio = False
End Function

Public Function fun825_ProcesarHojasEnvioCorregido(ByRef intHojasOcultadas As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR CORREGIDA: fun825_ProcesarHojasEnvioCorregido
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Fecha y Hora de Modificación: 2025-06-03 04:36:36 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Procesa hojas de envío, manteniendo visibles SOLO las 3 más recientes
    ' cuando hay 4 o más hojas de envío. Usa los últimos 16 caracteres para comparar fechas.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Variables para almacenar información de hojas
    Dim arrHojasEnvio() As String
    Dim arrSufijos() As String
    Dim i As Long, j As Long
    Dim intContador As Integer
    Dim strTemp As String
    Dim strTempSufijo As String
    Dim ws As Worksheet
    
    ' Inicializar contador
    intContador = 0
    intHojasOcultadas = 0
    
    '--------------------------------------------------------------------------
    ' 1. Recopilar todas las hojas de envío y sus sufijos
    '--------------------------------------------------------------------------
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        
        If fun824_EsHojaEnvio(ws.Name) Then
            intContador = intContador + 1
            ReDim Preserve arrHojasEnvio(1 To intContador)
            ReDim Preserve arrSufijos(1 To intContador)
            
            arrHojasEnvio(intContador) = ws.Name
            arrSufijos(intContador) = fun826_ExtraerSufijoCompleto(ws.Name)
            
            fun801_LogMessage "Hoja envío encontrada: " & ws.Name & " - Sufijo: " & arrSufijos(intContador), _
                              False, "", "fun825_ProcesarHojasEnvioCorregido"
        End If
    Next i
    
    '--------------------------------------------------------------------------
    ' 2. Ordenar por sufijo de mayor a menor (las más recientes primero)
    '--------------------------------------------------------------------------
    For i = 1 To intContador - 1
        For j = i + 1 To intContador
            ' Comparar sufijos: si el sufijo i es menor que j, intercambiar
            If arrSufijos(i) < arrSufijos(j) Then
                ' Intercambiar sufijos
                strTempSufijo = arrSufijos(i)
                arrSufijos(i) = arrSufijos(j)
                arrSufijos(j) = strTempSufijo
                
                ' Intercambiar nombres de hojas
                strTemp = arrHojasEnvio(i)
                arrHojasEnvio(i) = arrHojasEnvio(j)
                arrHojasEnvio(j) = strTemp
            End If
        Next j
    Next i
    
    '--------------------------------------------------------------------------
    ' 3. Log de orden final
    '--------------------------------------------------------------------------
    fun801_LogMessage "Orden final de hojas de envío (más recientes primero):", _
                      False, "", "fun825_ProcesarHojasEnvioCorregido"
    For i = 1 To intContador
        fun801_LogMessage "  " & i & ". " & arrHojasEnvio(i) & " - " & arrSufijos(i), _
                          False, "", "fun825_ProcesarHojasEnvioCorregido"
    Next i
    
    '--------------------------------------------------------------------------
    ' 4. Ocultar hojas desde la posición 4 en adelante (mantener solo las 3 primeras)
    '--------------------------------------------------------------------------
    For i = CONST_HOJAS_DE_ENVIO_VISIBLES + 1 To intContador
        Set ws = ThisWorkbook.Worksheets(arrHojasEnvio(i))
        If fun823_OcultarHojaSiVisible(ws) Then
            intHojasOcultadas = intHojasOcultadas + 1
            fun801_LogMessage "Hoja de envío ocultada: " & arrHojasEnvio(i), _
                              False, "", "fun825_ProcesarHojasEnvioCorregido"
        End If
    Next i
    
    '--------------------------------------------------------------------------
    ' 5. Log de hojas que permanecen visibles
    '--------------------------------------------------------------------------
    fun801_LogMessage "Hojas de envío que permanecen visibles:", _
                      False, "", "fun825_ProcesarHojasEnvioCorregido"
    For i = 1 To CONST_HOJAS_DE_ENVIO_VISIBLES
        If i <= intContador Then
            fun801_LogMessage "  Visible: " & arrHojasEnvio(i), _
                              False, "", "fun825_ProcesarHojasEnvioCorregido"
        End If
    Next i
    
    fun825_ProcesarHojasEnvioCorregido = True
    Exit Function
    
GestorErrores:
    fun801_LogMessage "Error en fun825_ProcesarHojasEnvioCorregido: " & Err.Description, _
                      True, "", "fun825_ProcesarHojasEnvioCorregido"
    fun825_ProcesarHojasEnvioCorregido = False
End Function

Public Function fun826_ExtraerSufijoCompleto(ByVal strNombreHoja As String) As String
    '******************************************************************************
    ' FUNCIÓN AUXILIAR CORREGIDA: fun826_ExtraerSufijoCompleto
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Fecha y Hora de Modificación: 2025-06-03 04:36:36 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Extrae los últimos 16 caracteres del nombre de hoja
    ' que contienen el sufijo de fecha/hora en formato _yyyyMMdd_hhmmss
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim strSufijo As String
    Dim intLongitudNombre As Integer
    
    intLongitudNombre = Len(strNombreHoja)
    
    ' Verificar que el nombre tenga al menos 16 caracteres
    If intLongitudNombre >= 16 Then
        ' Extraer los últimos 16 caracteres
        strSufijo = Right(strNombreHoja, 16)
        
        ' Verificar que tenga el formato correcto: _yyyyMMdd_hhmmss
        ' Debe tener un "_" en la posición 1 y otro "_" en la posición 10
        If Mid(strSufijo, 1, 1) = "_" And Mid(strSufijo, 10, 1) = "_" Then
            fun826_ExtraerSufijoCompleto = strSufijo
        Else
            ' Si no tiene el formato correcto, devolver un valor que lo coloque al final
            fun826_ExtraerSufijoCompleto = "_00000000_000000"
        End If
    Else
        ' Si no tiene suficiente longitud, devolver un valor que lo coloque al final
        fun826_ExtraerSufijoCompleto = "_00000000_000000"
    End If
    
    Exit Function
    
GestorErrores:
    fun826_ExtraerSufijoCompleto = "_00000000_000000"
End Function

Public Function fun821_ComenzarPorPrefijo(ByVal strTexto As String, ByVal strPrefijo As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun821_ComenzarPorPrefijo
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    If Len(strTexto) >= Len(strPrefijo) Then
        fun821_ComenzarPorPrefijo = (Left(strTexto, Len(strPrefijo)) = strPrefijo)
    Else
        fun821_ComenzarPorPrefijo = False
    End If
    Exit Function
    
ErrorHandler:
    fun821_ComenzarPorPrefijo = False
End Function

Public Function fun822_ValidarFormatoSufijoFecha(ByVal strNombreHoja As String, _
                                                ByVal strPrefijo As String, _
                                                ByVal intLongitudSufijo As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun822_ValidarFormatoSufijoFecha
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim intLongitudEsperada As Integer
    intLongitudEsperada = Len(strPrefijo) + intLongitudSufijo
    
    ' Validar longitud total
    If Len(strNombreHoja) = intLongitudEsperada Then
        fun822_ValidarFormatoSufijoFecha = True
    Else
        fun822_ValidarFormatoSufijoFecha = False
    End If
    Exit Function
    
ErrorHandler:
    fun822_ValidarFormatoSufijoFecha = False
End Function

Public Function fun823_ExtraerSufijoFecha(ByVal strNombreHoja As String, _
                                         ByVal intLongitudSufijo As Integer) As String
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_ExtraerSufijoFecha
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    If Len(strNombreHoja) >= intLongitudSufijo Then
        fun823_ExtraerSufijoFecha = Right(strNombreHoja, intLongitudSufijo)
    Else
        fun823_ExtraerSufijoFecha = ""
    End If
    Exit Function
    
ErrorHandler:
    fun823_ExtraerSufijoFecha = ""
End Function

Public Function fun824_CompararSufijosFecha(ByVal strSufijo1 As String, _
                                           ByVal strSufijo2 As String) As Integer
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun824_CompararSufijosFecha
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    ' Retorna: >0 si strSufijo1 > strSufijo2, 0 si iguales, <0 si strSufijo1 < strSufijo2
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    If strSufijo2 = "" Then
        fun824_CompararSufijosFecha = 1  ' strSufijo1 es mayor
    ElseIf strSufijo1 > strSufijo2 Then
        fun824_CompararSufijosFecha = 1
    ElseIf strSufijo1 < strSufijo2 Then
        fun824_CompararSufijosFecha = -1
    Else
        fun824_CompararSufijosFecha = 0
    End If
    Exit Function
    
ErrorHandler:
    fun824_CompararSufijosFecha = 0
End Function

Public Function fun825_CopiarHojaConNuevoNombre(ByVal strHojaOrigen As String, _
                                               ByVal strHojaDestino As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun825_CopiarHojaConNuevoNombre
    ' Fecha y Hora de Creación: 2025-06-03 06:00:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Crea una copia completa de una hoja de trabajo existente y le asigna
    ' un nuevo nombre. Maneja conflictos de nombres eliminando hojas existentes.
    '
    ' Pasos:
    ' 1. Validar que la hoja origen existe
    ' 2. Generar nombre de destino si no se proporciona
    ' 3. Eliminar hoja destino si ya existe
    ' 4. Copiar hoja origen con nuevo nombre
    ' 5. Verificar que la copia se creó correctamente
    '
    ' Parámetros:
    ' - strHojaOrigen: Nombre de la hoja a copiar
    ' - strHojaDestino: Nombre para la nueva hoja copiada
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim strNombreDestino As String
    
    ' Inicialización
    strFuncion = "fun825_CopiarHojaConNuevoNombre"
    fun825_CopiarHojaConNuevoNombre = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que la hoja origen existe
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If Len(Trim(strHojaOrigen)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 851, strFuncion, _
            "El nombre de la hoja origen está vacío"
    End If
    
    If Not fun802_SheetExists(strHojaOrigen) Then
        Err.Raise ERROR_BASE_IMPORT + 852, strFuncion, _
            "La hoja origen no existe: " & strHojaOrigen
    End If
    
    Set wsOrigen = ThisWorkbook.Worksheets(strHojaOrigen)
    
    '--------------------------------------------------------------------------
    ' 2. Preparar nombre de destino
    '--------------------------------------------------------------------------
    lngLineaError = 40
    If Len(Trim(strHojaDestino)) = 0 Then
        ' Generar nombre automático basado en timestamp
        strNombreDestino = strHojaOrigen & "_Copia_" & Format(Now(), "yyyymmdd_hhmmss")
    Else
        strNombreDestino = Trim(strHojaDestino)
    End If
    
    ' Validar longitud del nombre (Excel tiene límite de 31 caracteres)
    If Len(strNombreDestino) > 31 Then
        strNombreDestino = Left(strNombreDestino, 31)
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Eliminar hoja destino si ya existe
    '--------------------------------------------------------------------------
    lngLineaError = 50
    If fun802_SheetExists(strNombreDestino) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(strNombreDestino).Delete
        Application.DisplayAlerts = True
        
        fun801_LogMessage "Hoja existente eliminada: " & strNombreDestino, False, "", strFuncion
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Copiar hoja origen con nuevo nombre
    '--------------------------------------------------------------------------
    lngLineaError = 60
    Application.ScreenUpdating = False
    
    ' Copiar la hoja al final del libro
    wsOrigen.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ' Obtener referencia a la hoja recién copiada
    Set wsDestino = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ' Asignar nuevo nombre
    wsDestino.Name = strNombreDestino
    
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 5. Verificar que la copia se creó correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 70
    If Not fun802_SheetExists(strNombreDestino) Then
        Err.Raise ERROR_BASE_IMPORT + 853, strFuncion, _
            "Error al verificar la creación de la hoja copiada: " & strNombreDestino
    End If
    
    fun801_LogMessage "Hoja copiada exitosamente: " & strHojaOrigen & " ? " & strNombreDestino, _
                      False, "", strFuncion
    
    fun825_CopiarHojaConNuevoNombre = True
    Exit Function

GestorErrores:
    ' Restaurar configuración
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja Origen: " & strHojaOrigen & vbCrLf & _
                      "Hoja Destino: " & strHojaDestino
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun825_CopiarHojaConNuevoNombre = False
End Function
