Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_04"

Option Explicit

'******************************************************************************
' FUNCI�N PRINCIPAL: Modificar_Scenario_Year_Entity_en_hoja_PLAH
' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripci�n:
' Funci�n para modificar las dimensiones Scenario, Year y Entity en una hoja
' espec�fica de Excel, actualizando los valores en las filas correspondientes
' dentro de un rango de columnas determinado.
'
' RESUMEN EXHAUSTIVO DE PASOS:
' 1. Validaci�n de par�metros de entrada
' 2. Verificaci�n de existencia de la hoja objetivo
' 3. Obtenci�n de referencia a la hoja de trabajo
' 4. Configuraci�n del entorno para optimizar rendimiento
' 5. Validaci�n de rangos de filas y columnas
' 6. Recorrido de columnas desde vColumnaInicialHeaders hasta vColumnaFinalHeaders
' 7. Asignaci�n de valores en fila vFilaScenario con vScenario_xPL
' 8. Asignaci�n de valores en fila vFilaYear con vYear_xPL
' 9. Asignaci�n de valores en fila vFilaEntity con vEntity_xPL
' 10. Restauraci�n del entorno de Excel
' 11. Registro del resultado en el sistema de logging
'
' Compatibilidad: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
'******************************************************************************

Public Function Modificar_Scenario_Year_Entity_en_hoja_PLAH( _
    ByVal vReport_PL_AH_Name As String, _
    ByVal vFilaScenario As Integer, _
    ByVal vFilaYear As Integer, _
    ByVal vFilaEntity As Integer, _
    ByVal vColumnaInicialHeaders As Integer, _
    ByVal vColumnaFinalHeaders As Integer, _
    ByVal vScenario_xPL As String, _
    ByVal vYear_xPL As String, _
    ByVal vEntity_xPL As String) As Boolean

    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hoja de trabajo
    Dim wsDestino As Worksheet
    
    ' Variables para bucles
    Dim i As Integer
    
    ' Variables para optimizaci�n
    Dim blnScreenUpdating As Boolean
    Dim blnEnableEvents As Boolean
    Dim xlCalculationMode As Long
    
    ' Inicializaci�n
    strFuncion = "Modificar_Scenario_Year_Entity_en_hoja_PLAH" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Modificar_Scenario_Year_Entity_en_hoja_PLAH"
    Modificar_Scenario_Year_Entity_en_hoja_PLAH = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validaci�n de par�metros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando validaci�n de par�metros de entrada...", False, "", strFuncion
    
    ' Validar nombre de hoja
    If Not fun827_ValidarNombreHoja(vReport_PL_AH_Name) Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "Nombre de hoja no v�lido: " & vReport_PL_AH_Name
    End If
    
    ' Validar filas
    If Not fun828_ValidarParametrosFila(vFilaScenario, vFilaYear, vFilaEntity) Then
        Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
            "Par�metros de fila no v�lidos. Scenario: " & vFilaScenario & _
            ", Year: " & vFilaYear & ", Entity: " & vFilaEntity
    End If
    
    ' Validar columnas
    If Not fun829_ValidarParametrosColumna(vColumnaInicialHeaders, vColumnaFinalHeaders) Then
        Err.Raise ERROR_BASE_IMPORT + 803, strFuncion, _
            "Par�metros de columna no v�lidos. Inicial: " & vColumnaInicialHeaders & _
            ", Final: " & vColumnaFinalHeaders
    End If
    
    ' Validar valores a asignar
    If Not fun830_ValidarValoresAsignar(vScenario_xPL, vYear_xPL, vEntity_xPL) Then
        Err.Raise ERROR_BASE_IMPORT + 804, strFuncion, _
            "Valores a asignar no v�lidos"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Verificaci�n de existencia de la hoja objetivo
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Verificando existencia de hoja objetivo...", False, "", vReport_PL_AH_Name
    
    If Not fun802_SheetExists(vReport_PL_AH_Name) Then
        Err.Raise ERROR_BASE_IMPORT + 805, strFuncion, _
            "La hoja especificada no existe: " & vReport_PL_AH_Name
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Obtenci�n de referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Obteniendo referencia a la hoja de trabajo...", False, "", vReport_PL_AH_Name
    
    Set wsDestino = ThisWorkbook.Worksheets(vReport_PL_AH_Name)
    
    If wsDestino Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 806, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vReport_PL_AH_Name
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Configuraci�n del entorno para optimizar rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Configurando entorno para optimizaci�n...", False, "", vReport_PL_AH_Name
    
    ' Guardar configuraci�n actual
    blnScreenUpdating = Application.ScreenUpdating
    blnEnableEvents = Application.EnableEvents
    xlCalculationMode = Application.Calculation
    
    ' Configurar para optimizaci�n
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '--------------------------------------------------------------------------
    ' 5. Validaci�n de rangos en la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 90
    fun801_LogMessage "Validando rangos en la hoja de trabajo...", False, "", vReport_PL_AH_Name
    
    If Not fun831_ValidarRangosEnHoja(wsDestino, vFilaScenario, vFilaYear, vFilaEntity, _
                                      vColumnaInicialHeaders, vColumnaFinalHeaders) Then
        Err.Raise ERROR_BASE_IMPORT + 807, strFuncion, _
            "Los rangos especificados exceden los l�mites de la hoja"
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Recorrido de columnas y asignaci�n de valores
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Iniciando recorrido de columnas para asignaci�n de valores...", _
                      False, "", vReport_PL_AH_Name
    
    For i = vColumnaInicialHeaders To vColumnaFinalHeaders
        '----------------------------------------------------------------------
        ' 6.1. Asignaci�n de valor Scenario en fila correspondiente
        '----------------------------------------------------------------------
        lngLineaError = 110
        wsDestino.Cells(vFilaScenario, i).Value = vScenario_xPL
        
        '----------------------------------------------------------------------
        ' 6.2. Asignaci�n de valor Year en fila correspondiente
        '----------------------------------------------------------------------
        lngLineaError = 120
        wsDestino.Cells(vFilaYear, i).Value = vYear_xPL
        
        '----------------------------------------------------------------------
        ' 6.3. Asignaci�n de valor Entity en fila correspondiente
        '----------------------------------------------------------------------
        lngLineaError = 130
        wsDestino.Cells(vFilaEntity, i).Value = vEntity_xPL
    Next i
    
    '--------------------------------------------------------------------------
    ' 7. Restauraci�n del entorno de Excel
    '--------------------------------------------------------------------------
    lngLineaError = 140
    fun801_LogMessage "Restaurando configuraci�n del entorno...", False, "", vReport_PL_AH_Name
    
    Application.Calculation = xlCalculationMode
    Application.EnableEvents = blnEnableEvents
    Application.ScreenUpdating = blnScreenUpdating
    
    '--------------------------------------------------------------------------
    ' 8. Registro del resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 150
    fun801_LogMessage "Modificaci�n completada exitosamente. Columnas procesadas: " & _
                      (vColumnaFinalHeaders - vColumnaInicialHeaders + 1), _
                      False, "", vReport_PL_AH_Name
    
    Modificar_Scenario_Year_Entity_en_hoja_PLAH = True
    Exit Function

GestorErrores:
    ' Restaurar configuraci�n del entorno en caso de error
    On Error Resume Next
    Application.Calculation = xlCalculationMode
    Application.EnableEvents = blnEnableEvents
    Application.ScreenUpdating = blnScreenUpdating
    On Error GoTo 0
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja: " & vReport_PL_AH_Name & vbCrLf & _
                      "Par�metros: Scenario(" & vFilaScenario & "), Year(" & vFilaYear & _
                      "), Entity(" & vFilaEntity & "), Cols(" & vColumnaInicialHeaders & _
                      "-" & vColumnaFinalHeaders & ")"
    
    fun801_LogMessage strMensajeError, True, "", vReport_PL_AH_Name
    Modificar_Scenario_Year_Entity_en_hoja_PLAH = False
End Function

'******************************************************************************
' FUNCIONES AUXILIARES PARA VALIDACI�N
'******************************************************************************

Public Function fun827_ValidarNombreHoja(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun827_ValidarNombreHoja
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que el nombre de hoja sea v�lido y no est� vac�o
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun827_ValidarNombreHoja = False
    
    ' Verificar que no est� vac�o
    If Len(Trim(strNombreHoja)) = 0 Then
        Exit Function
    End If
    
    ' Verificar que no contenga caracteres no v�lidos para nombres de hoja
    If InStr(strNombreHoja, "[") > 0 Or InStr(strNombreHoja, "]") > 0 Or _
       InStr(strNombreHoja, ":") > 0 Or InStr(strNombreHoja, "*") > 0 Or _
       InStr(strNombreHoja, "?") > 0 Or InStr(strNombreHoja, "/") > 0 Or _
       InStr(strNombreHoja, "\") > 0 Then
        Exit Function
    End If
    
    ' Verificar longitud m�xima (31 caracteres para Excel)
    If Len(strNombreHoja) > 31 Then
        Exit Function
    End If
    
    fun827_ValidarNombreHoja = True
    Exit Function
    
ErrorHandler:
    fun827_ValidarNombreHoja = False
End Function

Public Function fun828_ValidarParametrosFila(ByVal vFilaScenario As Integer, _
                                            ByVal vFilaYear As Integer, _
                                            ByVal vFilaEntity As Integer) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun828_ValidarParametrosFila
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los par�metros de fila sean v�lidos
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun828_ValidarParametrosFila = False
    
    ' Verificar que sean valores positivos
    If vFilaScenario <= 0 Or vFilaYear <= 0 Or vFilaEntity <= 0 Then
        Exit Function
    End If
    
    ' Verificar que no excedan el l�mite m�ximo de Excel (compatible con Excel 97)
    If vFilaScenario > 65536 Or vFilaYear > 65536 Or vFilaEntity > 65536 Then
        Exit Function
    End If
    
    fun828_ValidarParametrosFila = True
    Exit Function
    
ErrorHandler:
    fun828_ValidarParametrosFila = False
End Function

Public Function fun829_ValidarParametrosColumna(ByVal vColumnaInicial As Integer, _
                                               ByVal vColumnaFinal As Integer) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun829_ValidarParametrosColumna
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los par�metros de columna sean v�lidos
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun829_ValidarParametrosColumna = False
    
    ' Verificar que sean valores positivos
    If vColumnaInicial <= 0 Or vColumnaFinal <= 0 Then
        Exit Function
    End If
    
    ' Verificar que la columna inicial sea menor o igual que la final
    If vColumnaInicial > vColumnaFinal Then
        Exit Function
    End If
    
    ' Verificar que no excedan el l�mite m�ximo de Excel (compatible con Excel 97: 256 columnas)
    If vColumnaInicial > 256 Or vColumnaFinal > 256 Then
        Exit Function
    End If
    
    fun829_ValidarParametrosColumna = True
    Exit Function
    
ErrorHandler:
    fun829_ValidarParametrosColumna = False
End Function

Public Function fun830_ValidarValoresAsignar(ByVal vScenario As String, _
                                            ByVal vYear As String, _
                                            ByVal vEntity As String) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun830_ValidarValoresAsignar
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los valores a asignar sean v�lidos (pueden estar vac�os)
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    ' En este caso, permitimos valores vac�os ya que podr�an ser v�lidos
    ' Solo verificamos que no sean Nothing (aunque al ser String esto no aplica)
    
    ' Verificar longitud m�xima razonable para evitar problemas de memoria
    If Len(vScenario) > 255 Or Len(vYear) > 255 Or Len(vEntity) > 255 Then
        fun830_ValidarValoresAsignar = False
        Exit Function
    End If
    
    fun830_ValidarValoresAsignar = True
    Exit Function
    
ErrorHandler:
    fun830_ValidarValoresAsignar = False
End Function

Public Function fun831_ValidarRangosEnHoja(ByRef ws As Worksheet, _
                                          ByVal vFilaScenario As Integer, _
                                          ByVal vFilaYear As Integer, _
                                          ByVal vFilaEntity As Integer, _
                                          ByVal vColumnaInicial As Integer, _
                                          ByVal vColumnaFinal As Integer) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun831_ValidarRangosEnHoja
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los rangos especificados existan en la hoja
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun831_ValidarRangosEnHoja = False
    
    ' Verificar que la hoja sea v�lida
    If ws Is Nothing Then
        Exit Function
    End If
    
    ' Verificar que las filas est�n dentro del rango de la hoja
    If vFilaScenario > ws.Rows.Count Or vFilaYear > ws.Rows.Count Or vFilaEntity > ws.Rows.Count Then
        Exit Function
    End If
    
    ' Verificar que las columnas est�n dentro del rango de la hoja
    If vColumnaInicial > ws.Columns.Count Or vColumnaFinal > ws.Columns.Count Then
        Exit Function
    End If
    
    ' Intentar acceder a las celdas para verificar que son accesibles
    On Error Resume Next
    Dim testValue As Variant
    testValue = ws.Cells(vFilaScenario, vColumnaInicial).Value
    testValue = ws.Cells(vFilaYear, vColumnaFinal).Value
    testValue = ws.Cells(vFilaEntity, vColumnaInicial).Value
    
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    fun831_ValidarRangosEnHoja = True
    Exit Function
    
ErrorHandler:
    fun831_ValidarRangosEnHoja = False
End Function

Public Function Convertir_RangoCellsCells_a_RangoCFCF(ByVal vFilaInicial As Integer, _
                                                      ByVal vFilaFinal As Integer, _
                                                      ByVal vColumnaInicial As Integer, _
                                                      ByVal vColumnaFinal As Integer) As String
    
    '******************************************************************************
    ' FUNCI�N: Convertir_RangoCellsCells_a_RangoCFCF
    ' FECHA Y HORA DE CREACI�N: 2025-06-15 11:29:40 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Convierte coordenadas num�ricas de filas y columnas a formato de rango de Excel
    ' est�ndar tipo "A5:P100". Funci�n auxiliar para generaci�n din�mica de rangos
    ' de celdas en operaciones de manipulaci�n de hojas de c�lculo.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables de control de errores y validaci�n
    ' 2. Validaci�n exhaustiva de par�metros de entrada (rangos v�lidos)
    ' 3. Verificaci�n de l�gica de coordenadas (inicial <= final)
    ' 4. Conversi�n de n�meros de columna a letras usando funci�n auxiliar
    ' 5. Construcci�n del string de rango en formato Excel est�ndar
    ' 6. Validaci�n del resultado generado antes del retorno
    ' 7. Logging de operaci�n para debugging y auditor�a
    ' 8. Retorno del string de rango formateado
    ' 9. Manejo exhaustivo de errores con informaci�n detallada
    ' 10. Limpieza de recursos y logging de errores en caso de fallo
    '
    ' PAR�METROS:
    ' - vFilaInicial (Integer): N�mero de fila inicial (debe ser >= 1)
    ' - vFilaFinal (Integer): N�mero de fila final (debe ser >= vFilaInicial)
    ' - vColumnaInicial (Integer): N�mero de columna inicial (debe ser >= 1)
    ' - vColumnaFinal (Integer): N�mero de columna final (debe ser >= vColumnaInicial)
    '
    ' RETORNA: String - Rango en formato Excel (ej: "A5:P100") o cadena vac�a si error
    '
    ' EJEMPLOS DE USO:
    ' Dim strRango As String
    ' strRango = Convertir_RangoCellsCells_a_RangoCFCF(5, 100, 1, 16)    ' Devuelve "A5:P100"
    ' strRango = Convertir_RangoCellsCells_a_RangoCFCF(1, 1, 1, 1)       ' Devuelve "A1:A1"
    ' strRango = Convertir_RangoCellsCells_a_RangoCFCF(10, 20, 5, 8)     ' Devuelve "E10:H20"
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento
    Dim strColumnaInicialLetra As String
    Dim strColumnaFinalLetra As String
    Dim strRangoResultado As String
    
    ' Inicializaci�n
    strFuncion = "Convertir_RangoCellsCells_a_RangoCFCF"
    Convertir_RangoCellsCells_a_RangoCFCF = ""
    lngLineaError = 0

    On Error GoTo GestorErrores

    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables de control de errores y validaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 30

    ' Inicializar variables de trabajo
    strColumnaInicialLetra = ""
    strColumnaFinalLetra = ""
    strRangoResultado = ""

    '--------------------------------------------------------------------------
    ' 2. Validaci�n exhaustiva de par�metros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 40

    ' Validar fila inicial
    If vFilaInicial < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9101, strFuncion, _
            "Fila inicial debe ser mayor que 0. Valor recibido: " & vFilaInicial
    End If

    ' Validar fila final
    If vFilaFinal < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9102, strFuncion, _
            "Fila final debe ser mayor que 0. Valor recibido: " & vFilaFinal
    End If

    ' Validar columna inicial
    If vColumnaInicial < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9103, strFuncion, _
            "Columna inicial debe ser mayor que 0. Valor recibido: " & vColumnaInicial
    End If

    ' Validar columna final
    If vColumnaFinal < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9104, strFuncion, _
            "Columna final debe ser mayor que 0. Valor recibido: " & vColumnaFinal
    End If

    ' Validar l�mites m�ximos de Excel (compatible con Excel 97-365)
    If vFilaInicial > 65536 Then
        Err.Raise ERROR_BASE_IMPORT + 9105, strFuncion, _
            "Fila inicial excede l�mite m�ximo de Excel (65536). Valor recibido: " & vFilaInicial
    End If

    If vFilaFinal > 65536 Then
        Err.Raise ERROR_BASE_IMPORT + 9106, strFuncion, _
            "Fila final excede l�mite m�ximo de Excel (65536). Valor recibido: " & vFilaFinal
    End If

    If vColumnaInicial > 256 Then
        Err.Raise ERROR_BASE_IMPORT + 9107, strFuncion, _
            "Columna inicial excede l�mite m�ximo de Excel (256). Valor recibido: " & vColumnaInicial
    End If

    If vColumnaFinal > 256 Then
        Err.Raise ERROR_BASE_IMPORT + 9108, strFuncion, _
            "Columna final excede l�mite m�ximo de Excel (256). Valor recibido: " & vColumnaFinal
    End If

    '--------------------------------------------------------------------------
    ' 3. Verificaci�n de l�gica de coordenadas
    '--------------------------------------------------------------------------
    lngLineaError = 50

    ' Verificar que fila inicial <= fila final
    If vFilaInicial > vFilaFinal Then
        Err.Raise ERROR_BASE_IMPORT + 9109, strFuncion, _
            "Fila inicial (" & vFilaInicial & ") debe ser menor o igual que fila final (" & vFilaFinal & ")"
    End If

    ' Verificar que columna inicial <= columna final
    If vColumnaInicial > vColumnaFinal Then
        Err.Raise ERROR_BASE_IMPORT + 9110, strFuncion, _
            "Columna inicial (" & vColumnaInicial & ") debe ser menor o igual que columna final (" & vColumnaFinal & ")"
    End If

    '--------------------------------------------------------------------------
    ' 4. Conversi�n de n�meros de columna a letras
    '--------------------------------------------------------------------------
    lngLineaError = 60

    ' Convertir columna inicial a letra usando funci�n auxiliar
    strColumnaInicialLetra = fun801_ConvertirNumeroColumnaALetra(vColumnaInicial)
    
    If Len(strColumnaInicialLetra) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9111, strFuncion, _
            "Error al convertir columna inicial a letra. Columna: " & vColumnaInicial
    End If

    ' Convertir columna final a letra usando funci�n auxiliar
    strColumnaFinalLetra = fun801_ConvertirNumeroColumnaALetra(vColumnaFinal)

    If Len(strColumnaFinalLetra) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9112, strFuncion, _
            "Error al convertir columna final a letra. Columna: " & vColumnaFinal
    End If

    '--------------------------------------------------------------------------
    ' 5. Construcci�n del string de rango en formato Excel est�ndar
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Construir el rango en formato "COLUMNA_INICIAL+FILA_INICIAL:COLUMNA_FINAL+FILA_FINAL"
    strRangoResultado = strColumnaInicialLetra & CStr(vFilaInicial) & Chr(58) & _
                        strColumnaFinalLetra & CStr(vFilaFinal)

    '--------------------------------------------------------------------------
    ' 6. Validaci�n del resultado generado
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Verificar que el resultado no est� vac�o
    If Len(strRangoResultado) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9113, strFuncion, _
            "Error al generar string de rango - resultado vac�o"
    End If
    
    ' Verificar que contiene el separador de rango (:)
    If InStr(strRangoResultado, Chr(58)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9114, strFuncion, _
            "Error en formato de rango - separador no encontrado: " & strRangoResultado
    End If
    
    ' Verificar longitud m�nima (ej: "A1:A1" = 5 caracteres m�nimo)
    If Len(strRangoResultado) < 5 Then
        Err.Raise ERROR_BASE_IMPORT + 9115, strFuncion, _
            "Longitud de rango inv�lida: " & strRangoResultado & " (Longitud: " & Len(strRangoResultado) & ")"
    End If

    '--------------------------------------------------------------------------
    ' 7. Logging de operaci�n para debugging y auditor�a
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    Call fun801_LogMessage("CONVERSI�N EXITOSA - Rango generado: " & Chr(34) & strRangoResultado & Chr(34) & _
        " desde coordenadas F(" & vFilaInicial & ":" & vFilaFinal & ") C(" & _
        vColumnaInicial & ":" & vColumnaFinal & ") = (" & strColumnaInicialLetra & ":" & _
        strColumnaFinalLetra & ")", False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 8. Retorno del string de rango formateado
    '--------------------------------------------------------------------------
    lngLineaError = 100
    Convertir_RangoCellsCells_a_RangoCFCF = strRangoResultado
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 9. Manejo exhaustivo de errores con informaci�n detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Par�metros de entrada:" & vbCrLf & _
                      "  - Fila inicial: " & vFilaInicial & vbCrLf & _
                      "  - Fila final: " & vFilaFinal & vbCrLf & _
                      "  - Columna inicial: " & vColumnaInicial & vbCrLf & _
                      "  - Columna final: " & vColumnaFinal & vbCrLf & _
                      "Variables de trabajo:" & vbCrLf & _
                      "  - Columna inicial letra: " & Chr(34) & strColumnaInicialLetra & Chr(34) & vbCrLf & _
                      "  - Columna final letra: " & Chr(34) & strColumnaFinalLetra & Chr(34) & vbCrLf & _
                      "  - Rango resultado: " & Chr(34) & strRangoResultado & Chr(34) & vbCrLf & _
                      "Fecha y Hora: " & Now() & vbCrLf & _
                      "Compatibilidad: Excel 97/2003/2007/365, OneDrive/SharePoint/Teams"

    '--------------------------------------------------------------------------
    ' 10. Logging de errores y limpieza de recursos
    '--------------------------------------------------------------------------
    
    ' Registrar error completo en log del sistema
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Para debugging en desarrollo
    Debug.Print strMensajeError
    
    ' Retornar cadena vac�a para indicar error
    Convertir_RangoCellsCells_a_RangoCFCF = ""
    
End Function

Public Function fun801_ConvertirNumeroColumnaALetra(ByVal vNumeroColumna As Integer) As String
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun801_ConvertirNumeroColumnaALetra
    ' FECHA Y HORA DE CREACI�N: 2025-06-15 11:29:40 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Convierte un n�mero de columna (1, 2, 3...) a su letra correspondiente
    ' en Excel (A, B, C, AA, AB...). Funci�n auxiliar para conversi�n de rangos.
    '
    ' PAR�METROS:
    ' - vNumeroColumna (Integer): N�mero de columna (1-256 para compatibilidad Excel 97)
    '
    ' RETORNA: String - Letra(s) de columna Excel o cadena vac�a si error
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim strResultado As String
    Dim intNumero As Integer
    
    ' Inicializaci�n
    fun801_ConvertirNumeroColumnaALetra = ""
    
    ' Validar par�metro
    If vNumeroColumna < 1 Or vNumeroColumna > 256 Then
        Exit Function
    End If
    
    ' Algoritmo de conversi�n a base 26 (letras A-Z)
    intNumero = vNumeroColumna
    strResultado = ""
    
    Do While intNumero > 0
        intNumero = intNumero - 1  ' Ajustar para base 0
        strResultado = Chr(65 + (intNumero Mod 26)) & strResultado
        intNumero = intNumero \ 26
    Loop
    
    fun801_ConvertirNumeroColumnaALetra = strResultado
    Exit Function
    
ErrorHandler:
    fun801_ConvertirNumeroColumnaALetra = ""
End Function

