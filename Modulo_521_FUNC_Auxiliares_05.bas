Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_05"
Option Explicit

Public Function Contiene_Scenario_Year_Entity(ByVal vSheet As String, _
                                             ByVal vEscenario As String, _
                                             ByVal vAnio As String, _
                                             ByVal vSociedad As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN: Contiene_Scenario_Year_Entity
    ' FECHA Y HORA DE CREACIÓN: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCIÓN:
    ' Función que verifica si una hoja específica contiene tres valores exactos:
    ' escenario, año y sociedad. Realiza búsqueda exhaustiva en toda la hoja
    ' para determinar si los tres valores están presentes como contenido completo
    ' de celdas individuales.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables de control de errores y optimización
    ' 2. Validación de parámetros de entrada y longitudes
    ' 3. Configuración de optimizaciones de rendimiento (pantalla, cálculo)
    ' 4. Verificación de existencia de la hoja especificada
    ' 5. Obtención de referencia a la hoja de trabajo
    ' 6. Determinación del rango usado para búsqueda eficiente
    ' 7. Búsqueda del primer valor (escenario) con coincidencia exacta
    ' 8. Búsqueda del segundo valor (año) con coincidencia exacta
    ' 9. Búsqueda del tercer valor (sociedad) con coincidencia exacta
    ' 10. Evaluación de resultados y determinación del valor de retorno
    ' 11. Registro de resultados en log del sistema
    ' 12. Restauración de configuraciones de optimización
    ' 13. Manejo exhaustivo de errores con información detallada
    '
    ' PARÁMETROS:
    ' - vSheet (String): Nombre de la hoja donde realizar la búsqueda
    ' - vEscenario (String): Valor del escenario a buscar (coincidencia exacta)
    ' - vAnio (String): Valor del año a buscar (coincidencia exacta)
    ' - vSociedad (String): Valor de la sociedad a buscar (coincidencia exacta)
    '
    ' VALOR DE RETORNO:
    ' - Boolean: True si los tres valores existen en la hoja, False en caso contrario
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    ' VERSIÓN: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimización
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas y búsqueda
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim blnHojaExiste As Boolean
    Dim rngUsedRange As Range
    
    ' Variables para resultados de búsqueda
    Dim blnExisteEscenario As Boolean
    Dim blnExisteAnio As Boolean
    Dim blnExisteSociedad As Boolean
    
    ' Variables para log de resultados
    Dim strMensajeLog As String
    
    ' Inicialización
    strFuncion = "Contiene_Scenario_Year_Entity"
    Contiene_Scenario_Year_Entity = False
    lngLineaError = 0
    blnExisteEscenario = False
    blnExisteAnio = False
    blnExisteSociedad = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables de control de errores y optimización
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando búsqueda de valores en hoja", False, "", vSheet
    
    ' Almacenar configuraciones originales para restaurar después
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Validación de parámetros de entrada y longitudes
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Validar que el nombre de la hoja no esté vacío
    If Len(Trim(vSheet)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "Parámetro vSheet está vacío"
    End If
    
    ' Validar que el escenario no esté vacío
    If Len(Trim(vEscenario)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
            "Parámetro vEscenario está vacío"
    End If
    
    ' Validar que el año no esté vacío
    If Len(Trim(vAnio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 803, strFuncion, _
            "Parámetro vAnio está vacío"
    End If
    
    ' Validar que la sociedad no esté vacía
    If Len(Trim(vSociedad)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 804, strFuncion, _
            "Parámetro vSociedad está vacío"
    End If
    
    ' Validar longitudes máximas razonables (compatibilidad Excel 97-365)
    If Len(Trim(vSheet)) > 31 Then
        Err.Raise ERROR_BASE_IMPORT + 805, strFuncion, _
            "Nombre de hoja demasiado largo: " & Len(Trim(vSheet)) & " caracteres"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Configuración de optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Desactivar actualización de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar cálculo automático para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 4. Verificación de existencia de la hoja especificada
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Obtener referencia al libro actual
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    ' Verificar que tenemos una referencia válida al libro
    If wb Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 806, strFuncion, _
            "No se pudo obtener referencia al libro de trabajo"
    End If
    
    ' Verificar existencia de la hoja usando función auxiliar existente del proyecto
    blnHojaExiste = fun801_VerificarExistenciaHoja(wb, vSheet)
    
    If Not blnHojaExiste Then
        Err.Raise ERROR_BASE_IMPORT + 807, strFuncion, _
            "La hoja especificada no existe: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Obtención de referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 90
    Set ws = wb.Worksheets(vSheet)
    
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 808, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Determinación del rango usado para búsqueda eficiente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Obtener rango usado de la hoja para optimizar búsqueda
    Set rngUsedRange = ws.UsedRange
    
    ' Verificar que la hoja tiene contenido
    If rngUsedRange Is Nothing Then
        fun801_LogMessage "Hoja está vacía, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    ' Verificar que el rango usado no está vacío
    If rngUsedRange.Cells.Count = 0 Then
        fun801_LogMessage "Rango usado está vacío, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Búsqueda del primer valor (escenario) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 110
    
    blnExisteEscenario = fun801_BuscarValorExactoEnRango(rngUsedRange, vEscenario)
    
    fun801_LogMessage "Búsqueda escenario " & Chr(34) & vEscenario & Chr(34) & _
        " resultado: " & blnExisteEscenario, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 8. Búsqueda del segundo valor (año) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
    blnExisteAnio = fun801_BuscarValorExactoEnRango(rngUsedRange, vAnio)
    
    fun801_LogMessage "Búsqueda año " & Chr(34) & vAnio & Chr(34) & _
        " resultado: " & blnExisteAnio, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 9. Búsqueda del tercer valor (sociedad) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 130
    
    blnExisteSociedad = fun801_BuscarValorExactoEnRango(rngUsedRange, vSociedad)
    
    fun801_LogMessage "Búsqueda sociedad " & Chr(34) & vSociedad & Chr(34) & _
        " resultado: " & blnExisteSociedad, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 10. Evaluación de resultados y determinación del valor de retorno
    '--------------------------------------------------------------------------
    lngLineaError = 140
    
    ' La función retorna True solo si los tres valores existen
    If blnExisteEscenario And blnExisteAnio And blnExisteSociedad Then
        Contiene_Scenario_Year_Entity = True
        strMensajeLog = "ÉXITO - Los tres valores existen en la hoja"
    Else
        Contiene_Scenario_Year_Entity = False
        strMensajeLog = "RESULTADO - Valores faltantes: "
        If Not blnExisteEscenario Then strMensajeLog = strMensajeLog & "Escenario "
        If Not blnExisteAnio Then strMensajeLog = strMensajeLog & "Año "
        If Not blnExisteSociedad Then strMensajeLog = strMensajeLog & "Sociedad "
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Registro de resultados en log del sistema
    '--------------------------------------------------------------------------
    lngLineaError = 150
    
    fun801_LogMessage strMensajeLog & " - Hoja: " & vSheet & _
        ", Escenario: " & Chr(34) & vEscenario & Chr(34) & _
        ", Año: " & Chr(34) & vAnio & Chr(34) & _
        ", Sociedad: " & Chr(34) & vSociedad & Chr(34) & _
        ", Resultado final: " & Contiene_Scenario_Year_Entity, _
        False, "", vSheet

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 12. Restauración de configuraciones de optimización
    '--------------------------------------------------------------------------
    lngLineaError = 160
    
    ' Restaurar configuración original de actualización de pantalla
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Restaurar configuración original de cálculo
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    ' Restaurar configuración original de eventos
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set rngUsedRange = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    fun801_LogMessage "Búsqueda completada exitosamente", False, "", vSheet
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 13. Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja: " & vSheet & vbCrLf & _
                      "Escenario: " & Chr(34) & vEscenario & Chr(34) & vbCrLf & _
                      "Año: " & Chr(34) & vAnio & Chr(34) & vbCrLf & _
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
    ' FUNCIÓN AUXILIAR: fun801_BuscarValorExactoEnRango
    ' FECHA Y HORA DE CREACIÓN: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Busca un valor específico dentro de un rango de celdas con coincidencia exacta
    ' y comparación case-insensitive. Optimizada para compatibilidad Excel 97-365.
    '
    ' PARÁMETROS:
    ' - rngBusqueda (Range): Rango donde realizar la búsqueda
    ' - strValorBuscado (String): Valor a buscar con coincidencia exacta
    '
    ' RETORNA: Boolean - True si encuentra el valor, False si no lo encuentra
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Variables para búsqueda
    Dim rngCelda As Range
    Dim rngEncontrado As Range
    Dim strValorCelda As String
    Dim strValorBuscadoNormalizado As String
    
    ' Inicialización
    strFuncion = "fun801_BuscarValorExactoEnRango"
    fun801_BuscarValorExactoEnRango = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validación de parámetros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    If rngBusqueda Is Nothing Then
        Exit Function
    End If
    
    If Len(Trim(strValorBuscado)) = 0 Then
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Normalización del valor buscado para comparación case-insensitive
    '--------------------------------------------------------------------------
    lngLineaError = 40
    strValorBuscadoNormalizado = UCase(Trim(strValorBuscado))
    
    '--------------------------------------------------------------------------
    ' 3. Búsqueda usando método Find (más eficiente para rangos grandes)
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Usar Find con configuración compatible Excel 97-365
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
    ' 4. Método alternativo: búsqueda manual (fallback para casos especiales)
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Si Find no funcionó, usar método manual como respaldo
    For Each rngCelda In rngBusqueda.Cells
        ' Verificar que la celda no esté vacía
        If Not IsEmpty(rngCelda.Value) And Not IsNull(rngCelda.Value) Then
            strValorCelda = UCase(Trim(CStr(rngCelda.Value)))
            
            ' Comparación exacta case-insensitive
            If strValorCelda = strValorBuscadoNormalizado Then
                fun801_BuscarValorExactoEnRango = True
                Exit Function
            End If
        End If
    Next rngCelda
    
    ' Si llegamos aquí, no se encontró el valor
    fun801_BuscarValorExactoEnRango = False
    Exit Function

GestorErrores:
    ' En caso de error, retornar False
    fun801_BuscarValorExactoEnRango = False
    
    ' Log del error para debugging
    Debug.Print "Error en " & strFuncion & " línea " & lngLineaError & ": " & Err.Description
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

Public Function fun823_MostrarHojaSiOculta(vNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_MostrarHojaSiOculta
    ' Fecha y Hora de Creación: 2025-06-08 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Muestra una hoja si está oculta
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


