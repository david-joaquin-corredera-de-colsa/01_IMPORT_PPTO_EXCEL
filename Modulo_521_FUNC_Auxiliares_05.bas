Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_05"
Option Explicit

Public Function Contiene_Scenario_Year_Entity(ByVal vSheet As String, _
                                             ByVal vEscenario As String, _
                                             ByVal vAnio As String, _
                                             ByVal vSociedad As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: Contiene_Scenario_Year_Entity
    ' FECHA Y HORA DE CREACI�N: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Funci�n que verifica si una hoja espec�fica contiene tres valores exactos:
    ' escenario, a�o y sociedad. Realiza b�squeda exhaustiva en toda la hoja
    ' para determinar si los tres valores est�n presentes como contenido completo
    ' de celdas individuales.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables de control de errores y optimizaci�n
    ' 2. Validaci�n de par�metros de entrada y longitudes
    ' 3. Configuraci�n de optimizaciones de rendimiento (pantalla, c�lculo)
    ' 4. Verificaci�n de existencia de la hoja especificada
    ' 5. Obtenci�n de referencia a la hoja de trabajo
    ' 6. Determinaci�n del rango usado para b�squeda eficiente
    ' 7. B�squeda del primer valor (escenario) con coincidencia exacta
    ' 8. B�squeda del segundo valor (a�o) con coincidencia exacta
    ' 9. B�squeda del tercer valor (sociedad) con coincidencia exacta
    ' 10. Evaluaci�n de resultados y determinaci�n del valor de retorno
    ' 11. Registro de resultados en log del sistema
    ' 12. Restauraci�n de configuraciones de optimizaci�n
    ' 13. Manejo exhaustivo de errores con informaci�n detallada
    '
    ' PAR�METROS:
    ' - vSheet (String): Nombre de la hoja donde realizar la b�squeda
    ' - vEscenario (String): Valor del escenario a buscar (coincidencia exacta)
    ' - vAnio (String): Valor del a�o a buscar (coincidencia exacta)
    ' - vSociedad (String): Valor de la sociedad a buscar (coincidencia exacta)
    '
    ' VALOR DE RETORNO:
    ' - Boolean: True si los tres valores existen en la hoja, False en caso contrario
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    ' VERSI�N: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimizaci�n
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas y b�squeda
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim blnHojaExiste As Boolean
    Dim rngUsedRange As Range
    
    ' Variables para resultados de b�squeda
    Dim blnExisteEscenario As Boolean
    Dim blnExisteAnio As Boolean
    Dim blnExisteSociedad As Boolean
    
    ' Variables para log de resultados
    Dim strMensajeLog As String
    
    ' Inicializaci�n
    strFuncion = "Contiene_Scenario_Year_Entity"
    Contiene_Scenario_Year_Entity = False
    lngLineaError = 0
    blnExisteEscenario = False
    blnExisteAnio = False
    blnExisteSociedad = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables de control de errores y optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando b�squeda de valores en hoja", False, "", vSheet
    
    ' Almacenar configuraciones originales para restaurar despu�s
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Validaci�n de par�metros de entrada y longitudes
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Validar que el nombre de la hoja no est� vac�o
    If Len(Trim(vSheet)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "Par�metro vSheet est� vac�o"
    End If
    
    ' Validar que el escenario no est� vac�o
    If Len(Trim(vEscenario)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
            "Par�metro vEscenario est� vac�o"
    End If
    
    ' Validar que el a�o no est� vac�o
    If Len(Trim(vAnio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 803, strFuncion, _
            "Par�metro vAnio est� vac�o"
    End If
    
    ' Validar que la sociedad no est� vac�a
    If Len(Trim(vSociedad)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 804, strFuncion, _
            "Par�metro vSociedad est� vac�o"
    End If
    
    ' Validar longitudes m�ximas razonables (compatibilidad Excel 97-365)
    If Len(Trim(vSheet)) > 31 Then
        Err.Raise ERROR_BASE_IMPORT + 805, strFuncion, _
            "Nombre de hoja demasiado largo: " & Len(Trim(vSheet)) & " caracteres"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Configuraci�n de optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Desactivar actualizaci�n de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar c�lculo autom�tico para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 4. Verificaci�n de existencia de la hoja especificada
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Obtener referencia al libro actual
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    ' Verificar que tenemos una referencia v�lida al libro
    If wb Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 806, strFuncion, _
            "No se pudo obtener referencia al libro de trabajo"
    End If
    
    ' Verificar existencia de la hoja usando funci�n auxiliar existente del proyecto
    blnHojaExiste = fun801_VerificarExistenciaHoja(wb, vSheet)
    
    If Not blnHojaExiste Then
        Err.Raise ERROR_BASE_IMPORT + 807, strFuncion, _
            "La hoja especificada no existe: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Obtenci�n de referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 90
    Set ws = wb.Worksheets(vSheet)
    
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 808, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Determinaci�n del rango usado para b�squeda eficiente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Obtener rango usado de la hoja para optimizar b�squeda
    Set rngUsedRange = ws.UsedRange
    
    ' Verificar que la hoja tiene contenido
    If rngUsedRange Is Nothing Then
        fun801_LogMessage "Hoja est� vac�a, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    ' Verificar que el rango usado no est� vac�o
    If rngUsedRange.Cells.Count = 0 Then
        fun801_LogMessage "Rango usado est� vac�o, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    '--------------------------------------------------------------------------
    ' 7. B�squeda del primer valor (escenario) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 110
    
    blnExisteEscenario = fun801_BuscarValorExactoEnRango(rngUsedRange, vEscenario)
    
    fun801_LogMessage "B�squeda escenario " & Chr(34) & vEscenario & Chr(34) & _
        " resultado: " & blnExisteEscenario, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 8. B�squeda del segundo valor (a�o) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
    blnExisteAnio = fun801_BuscarValorExactoEnRango(rngUsedRange, vAnio)
    
    fun801_LogMessage "B�squeda a�o " & Chr(34) & vAnio & Chr(34) & _
        " resultado: " & blnExisteAnio, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 9. B�squeda del tercer valor (sociedad) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 130
    
    blnExisteSociedad = fun801_BuscarValorExactoEnRango(rngUsedRange, vSociedad)
    
    fun801_LogMessage "B�squeda sociedad " & Chr(34) & vSociedad & Chr(34) & _
        " resultado: " & blnExisteSociedad, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 10. Evaluaci�n de resultados y determinaci�n del valor de retorno
    '--------------------------------------------------------------------------
    lngLineaError = 140
    
    ' La funci�n retorna True solo si los tres valores existen
    If blnExisteEscenario And blnExisteAnio And blnExisteSociedad Then
        Contiene_Scenario_Year_Entity = True
        strMensajeLog = "�XITO - Los tres valores existen en la hoja"
    Else
        Contiene_Scenario_Year_Entity = False
        strMensajeLog = "RESULTADO - Valores faltantes: "
        If Not blnExisteEscenario Then strMensajeLog = strMensajeLog & "Escenario "
        If Not blnExisteAnio Then strMensajeLog = strMensajeLog & "A�o "
        If Not blnExisteSociedad Then strMensajeLog = strMensajeLog & "Sociedad "
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Registro de resultados en log del sistema
    '--------------------------------------------------------------------------
    lngLineaError = 150
    
    fun801_LogMessage strMensajeLog & " - Hoja: " & vSheet & _
        ", Escenario: " & Chr(34) & vEscenario & Chr(34) & _
        ", A�o: " & Chr(34) & vAnio & Chr(34) & _
        ", Sociedad: " & Chr(34) & vSociedad & Chr(34) & _
        ", Resultado final: " & Contiene_Scenario_Year_Entity, _
        False, "", vSheet

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 12. Restauraci�n de configuraciones de optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 160
    
    ' Restaurar configuraci�n original de actualizaci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Restaurar configuraci�n original de c�lculo
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    ' Restaurar configuraci�n original de eventos
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set rngUsedRange = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    fun801_LogMessage "B�squeda completada exitosamente", False, "", vSheet
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 13. Manejo exhaustivo de errores con informaci�n detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja: " & vSheet & vbCrLf & _
                      "Escenario: " & Chr(34) & vEscenario & Chr(34) & vbCrLf & _
                      "A�o: " & Chr(34) & vAnio & Chr(34) & vbCrLf & _
                      "Sociedad: " & Chr(34) & vSociedad & Chr(34) & vbCrLf & _
                      "Fecha y Hora: " & Now()
    
    ' Registrar error en log del sistema
    fun801_LogMessage strMensajeError, True, "", vSheet
    
    ' Log del error para debugging
    Debug.Print strMensajeError
    
    ' Restaurar configuraciones en caso de error
    On Error Resume Next
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set rngUsedRange = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    ' Retornar False para indicar error
    Contiene_Scenario_Year_Entity = False
End Function

Public Function fun801_BuscarValorExactoEnRango(ByRef rngBusqueda As Range, _
                                               ByVal strValorBuscado As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun801_BuscarValorExactoEnRango
    ' FECHA Y HORA DE CREACI�N: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Busca un valor espec�fico dentro de un rango de celdas con coincidencia exacta
    ' y comparaci�n case-insensitive. Optimizada para compatibilidad Excel 97-365.
    '
    ' PAR�METROS:
    ' - rngBusqueda (Range): Rango donde realizar la b�squeda
    ' - strValorBuscado (String): Valor a buscar con coincidencia exacta
    '
    ' RETORNA: Boolean - True si encuentra el valor, False si no lo encuentra
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Variables para b�squeda
    Dim rngCelda As Range
    Dim rngEncontrado As Range
    Dim strValorCelda As String
    Dim strValorBuscadoNormalizado As String
    
    ' Inicializaci�n
    strFuncion = "fun801_BuscarValorExactoEnRango"
    fun801_BuscarValorExactoEnRango = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validaci�n de par�metros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    If rngBusqueda Is Nothing Then
        Exit Function
    End If
    
    If Len(Trim(strValorBuscado)) = 0 Then
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Normalizaci�n del valor buscado para comparaci�n case-insensitive
    '--------------------------------------------------------------------------
    lngLineaError = 40
    strValorBuscadoNormalizado = UCase(Trim(strValorBuscado))
    
    '--------------------------------------------------------------------------
    ' 3. B�squeda usando m�todo Find (m�s eficiente para rangos grandes)
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Usar Find con configuraci�n compatible Excel 97-365
    Set rngEncontrado = rngBusqueda.Find( _
        What:=strValorBuscado, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    
    ' Si Find encuentra algo, verificar que sea coincidencia exacta
    If Not rngEncontrado Is Nothing Then
        strValorCelda = UCase(Trim(CStr(rngEncontrado.Value)))
        If strValorCelda = strValorBuscadoNormalizado Then
            fun801_BuscarValorExactoEnRango = True
            Exit Function
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 4. M�todo alternativo: b�squeda manual (fallback para casos especiales)
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Si Find no funcion�, usar m�todo manual como respaldo
    For Each rngCelda In rngBusqueda.Cells
        ' Verificar que la celda no est� vac�a
        If Not IsEmpty(rngCelda.Value) And Not IsNull(rngCelda.Value) Then
            strValorCelda = UCase(Trim(CStr(rngCelda.Value)))
            
            ' Comparaci�n exacta case-insensitive
            If strValorCelda = strValorBuscadoNormalizado Then
                fun801_BuscarValorExactoEnRango = True
                Exit Function
            End If
        End If
    Next rngCelda
    
    ' Si llegamos aqu�, no se encontr� el valor
    fun801_BuscarValorExactoEnRango = False
    Exit Function

GestorErrores:
    ' En caso de error, retornar False
    fun801_BuscarValorExactoEnRango = False
    
    ' Log del error para debugging
    Debug.Print "Error en " & strFuncion & " l�nea " & lngLineaError & ": " & Err.Description
End Function


Public Function fun824_LimpiarFilasExcedentes(ByRef ws As Worksheet, _
                                             ByVal vFila_Final_Limite As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun824_LimpiarFilasExcedentes
    ' Fecha y Hora de Creaci�n: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Limpia todas las filas que est�n por encima del l�mite especificado
    ' Borra tanto contenido como formatos para optimizar el archivo
    '
    ' Par�metros:
    ' - ws: Hoja de trabajo donde limpiar
    ' - vFila_Final_Limite: N�mero de fila l�mite (se borran filas superiores)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngUltimaFilaConDatos As Long
    
    ' Validar par�metros
    If ws Is Nothing Then
        fun824_LimpiarFilasExcedentes = False
        Exit Function
    End If
    
    If vFila_Final_Limite < 1 Then
        fun824_LimpiarFilasExcedentes = False
        Exit Function
    End If
    
    ' Obtener �ltima fila con datos (m�todo compatible Excel 97-365)
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
    ' FUNCI�N AUXILIAR: fun825_LimpiarColumnasExcedentes
    ' Fecha y Hora de Creaci�n: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Limpia todas las columnas que est�n por encima del l�mite especificado
    ' Borra tanto contenido como formatos para optimizar el archivo
    '
    ' Par�metros:
    ' - ws: Hoja de trabajo donde limpiar
    ' - vColumna_Final_Limite: N�mero de columna l�mite (se borran columnas superiores)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngUltimaColumnaConDatos As Long
    
    ' Validar par�metros
    If ws Is Nothing Then
        fun825_LimpiarColumnasExcedentes = False
        Exit Function
    End If
    
    If vColumna_Final_Limite < 1 Then
        fun825_LimpiarColumnasExcedentes = False
        Exit Function
    End If
    
    ' Obtener �ltima columna con datos (m�todo compatible Excel 97-365)
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
    ' FUNCI�N AUXILIAR: fun826_ConfigurarPalabrasClave
    ' Fecha y Hora de Creaci�n: 2025-06-03 03:19:45 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Permite configurar las palabras clave utilizadas para detectar rangos
    ' de datos en las hojas de trabajo.
    '
    ' Par�metros (todos opcionales):
    ' - strPrimeraFila: Palabra clave para buscar primera fila
    ' - strPrimeraColumna: Palabra clave para buscar primera columna
    ' - strUltimaFila: Palabra clave para buscar �ltima fila
    ' - strUltimaColumna: Palabra clave para buscar �ltima columna
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

Public Function fun823_OcultarHojaSiVisible(ByRef ws As Worksheet) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun823_OcultarHojaSiVisible
    ' Fecha y Hora de Creaci�n: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Oculta una hoja si est� visible
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

Public Function fun823_MostrarHojaSiOculta(vNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun823_MostrarHojaSiOculta
    ' Fecha y Hora de Creaci�n: 2025-06-08 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Muestra una hoja si est� oculta
    '******************************************************************************
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(vNombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then
        fun823_MostrarHojaSiOculta = False
    Else
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
        fun823_MostrarHojaSiOculta = True
    End If
End Function



Public Function fun821_ComenzarPorPrefijo(ByVal strTexto As String, ByVal strPrefijo As String) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun821_ComenzarPorPrefijo
    ' Fecha y Hora de Creaci�n: 2025-06-03 05:34:14 UTC
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
    ' FUNCI�N AUXILIAR: fun822_ValidarFormatoSufijoFecha
    ' Fecha y Hora de Creaci�n: 2025-06-03 05:34:14 UTC
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
    ' FUNCI�N AUXILIAR: fun823_ExtraerSufijoFecha
    ' Fecha y Hora de Creaci�n: 2025-06-03 05:34:14 UTC
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
    ' FUNCI�N AUXILIAR: fun824_CompararSufijosFecha
    ' Fecha y Hora de Creaci�n: 2025-06-03 05:34:14 UTC
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
    ' FUNCI�N AUXILIAR: fun825_CopiarHojaConNuevoNombre
    ' Fecha y Hora de Creaci�n: 2025-06-03 06:00:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Crea una copia completa de una hoja de trabajo existente y le asigna
    ' un nuevo nombre. Maneja conflictos de nombres eliminando hojas existentes.
    '
    ' Pasos:
    ' 1. Validar que la hoja origen existe
    ' 2. Generar nombre de destino si no se proporciona
    ' 3. Eliminar hoja destino si ya existe
    ' 4. Copiar hoja origen con nuevo nombre
    ' 5. Verificar que la copia se cre� correctamente
    '
    ' Par�metros:
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
    
    ' Inicializaci�n
    strFuncion = "fun825_CopiarHojaConNuevoNombre" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun825_CopiarHojaConNuevoNombre"
    fun825_CopiarHojaConNuevoNombre = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que la hoja origen existe
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If Len(Trim(strHojaOrigen)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 851, strFuncion, _
            "El nombre de la hoja origen est� vac�o"
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
        ' Generar nombre autom�tico basado en timestamp
        strNombreDestino = strHojaOrigen & "_Copia_" & Format(Now(), "yyyymmdd_hhmmss")
    Else
        strNombreDestino = Trim(strHojaDestino)
    End If
    
    ' Validar longitud del nombre (Excel tiene l�mite de 31 caracteres)
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
    
    ' Obtener referencia a la hoja reci�n copiada
    Set wsDestino = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ' Asignar nuevo nombre
    wsDestino.Name = strNombreDestino
    
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 5. Verificar que la copia se cre� correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 70
    If Not fun802_SheetExists(strNombreDestino) Then
        Err.Raise ERROR_BASE_IMPORT + 853, strFuncion, _
            "Error al verificar la creaci�n de la hoja copiada: " & strNombreDestino
    End If
    
    fun801_LogMessage "Hoja copiada exitosamente: " & strHojaOrigen & " ? " & strNombreDestino, _
                      False, "", strFuncion
    
    fun825_CopiarHojaConNuevoNombre = True
    Exit Function

GestorErrores:
    ' Restaurar configuraci�n
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja Origen: " & strHojaOrigen & vbCrLf & _
                      "Hoja Destino: " & strHojaDestino
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun825_CopiarHojaConNuevoNombre = False
End Function


