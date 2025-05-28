Attribute VB_Name = "Modulo_014_FUNC_Ppales_StoDelim"
' =============================================================================
' VARIABLES GLOBALES PARA DELIMITADORES DE EXCEL
' =============================================================================
' Fecha y hora de creación: 2025-05-26 17:43:59 UTC
' Autor: david-joaquin-corredera-de-colsa
' Descripción: Variables globales para el manejo de delimitadores de Excel
' =============================================================================

Public vHojaDelimitadoresExcelOriginales As String
Public vCelda_Header_Excel_UseSystemSeparators As String
Public vCelda_Header_Excel_DecimalSeparator As String
Public vCelda_Header_Excel_ThousandsSeparator As String
Public vCelda_Valor_Excel_UseSystemSeparators As String
Public vCelda_Valor_Excel_DecimalSeparator As String
Public vCelda_Valor_Excel_ThousandsSeparator As String
Public vExcel_UseSystemSeparators As String
Public vExcel_DecimalSeparator As String
Public vExcel_ThousandsSeparator As String
'Public vOcultarRepostiorioDelimitadores As Boolean
'vOcultarRepostiorioDelimitadores = True ' Cambiar a True si se desea ocultar la hoja
Public Const vOcultarRepostiorioDelimitadores As Boolean = True


' =============================================================================
' FUNCIÓN PRINCIPAL: F004_Detectar_Delimitadores_en_Excel
' =============================================================================
' Fecha y hora de creación: 2025-05-26 17:43:59 UTC
' Autor: david-joaquin-corredera-de-colsa
' Descripción: Detecta y almacena los delimitadores de Excel actuales
'
' RESUMEN EXHAUSTIVO DE PASOS:
' 1. Inicializar variables globales con valores por defecto
' 2. Verificar si existe la hoja de delimitadores originales
' 3. Si no existe, crear la hoja y dejarla visible
' 4. Si existe, verificar su visibilidad y hacerla visible si está oculta
' 5. Limpiar el contenido de la hoja una vez visible
' 6. Configurar headers en las celdas especificadas (B2, B3, B4)
' 7. Detectar configuración actual de delimitadores de Excel:
'    - Use System Separators (True/False)
'    - Decimal Separator (carácter)
'    - Thousands Separator (carácter)
' 8. Almacenar valores detectados en variables globales
' 9. Escribir valores en la hoja de delimitadores (C2, C3, C4)
' 10. Verificar variable global vOcultarRepostiorioDelimitadores
' 11. Si es True, ocultar la hoja creada/actualizada
' 12. Manejo exhaustivo de errores con información detallada
'
' Parámetros: Ninguno
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Function F004_Detectar_Delimitadores_en_Excel() As Boolean
    
    ' Control de errores con número de línea
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
    
    ' Nombre de la hoja donde se almacenarán los delimitadores originales
    vHojaDelimitadoresExcelOriginales = "06_Delimitadores_Originales"
    
    ' Celdas para los headers (títulos)
    vCelda_Header_Excel_UseSystemSeparators = "B2"
    vCelda_Header_Excel_DecimalSeparator = "B3"
    vCelda_Header_Excel_ThousandsSeparator = "B4"
    
    ' Celdas para los valores detectados
    vCelda_Valor_Excel_UseSystemSeparators = "C2"
    vCelda_Valor_Excel_DecimalSeparator = "C3"
    vCelda_Valor_Excel_ThousandsSeparator = "C4"
    
    ' Variables para almacenar los valores detectados (inicialmente vacías)
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
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEGÚN CORRESPONDA
    ' ==========================================================================
    
    If Not hojaExiste Then
        ' La hoja no existe, crearla y dejarla visible
        Set ws = fun802_CrearHojaDelimitadores(wb, vHojaDelimitadoresExcelOriginales)
        If ws Is Nothing Then
            F004_Detectar_Delimitadores_en_Excel = False
            Exit Function
        End If
        ' La hoja recién creada ya está visible por defecto
    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
        Set ws = wb.Worksheets(vHojaDelimitadoresExcelOriginales)
        
        ' Verificar si está oculta y hacerla visible si es necesario
        Call fun803_HacerHojaVisible(ws)
    End If
    
    lineaError = 140
    
    ' ==========================================================================
    ' PASO 5: LIMPIAR CONTENIDO DE LA HOJA (AHORA QUE ESTÁ VISIBLE)
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
    ' PASO 7: DETECTAR CONFIGURACIÓN ACTUAL DE DELIMITADORES DE EXCEL
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
        Call fun808_OcultarHojaDelimitadores(ws)
    End If
    
    lineaError = 190
    
    ' ==========================================================================
    ' PASO 10: FINALIZACIÓN EXITOSA
    ' ==========================================================================
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Detectar_Delimitadores_en_Excel = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: F004_Detectar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now() & vbCrLf & _
                   "USUARIO: david-joaquin-corredera-de-colsa"
    
    ' Mostrar mensaje de error (comentar si no se desea)
    ' MsgBox mensajeError, vbCritical, "Error en Detección de Delimitadores"
    
    ' Log del error para debugging
    Debug.Print mensajeError
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 801: VERIFICAR EXISTENCIA DE HOJA
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Verifica si una hoja existe en el libro especificado
' Parámetros: wb (Workbook), nombreHoja (String)
' Retorna: Boolean (True si existe, False si no existe)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun801_VerificarExistenciaHoja(wb As Workbook, nombreHoja As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim lineaError As Long
    
    lineaError = 200
    fun801_VerificarExistenciaHoja = False
    
    ' Verificar parámetros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Exit Function
    End If
    
    lineaError = 210
    
    ' Recorrer todas las hojas del libro (método compatible con Excel 97)
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
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun801_VerificarExistenciaHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "PARÁMETRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 802: CREAR HOJA DE DELIMITADORES
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Crea una nueva hoja con el nombre especificado y la deja visible
' Parámetros: wb (Workbook), nombreHoja (String)
' Retorna: Worksheet (referencia a la hoja creada, Nothing si error)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Function fun802_CrearHojaDelimitadores(wb As Workbook, nombreHoja As String) As Worksheet
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lineaError As Long
    
    lineaError = 300
    
    ' Verificar parámetros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Exit Function
    End If
    
    lineaError = 310
    
    ' Verificar que el libro no esté protegido (importante para entornos cloud)
    If wb.ProtectStructure Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Exit Function
    End If
    
    lineaError = 320
    
    ' Crear nueva hoja al final del libro (método compatible con todas las versiones)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    lineaError = 330
    
    ' Asignar nombre a la hoja
    ws.Name = nombreHoja
    
    lineaError = 340
    
    ' Asegurar que la hoja esté visible (por defecto ya lo está, pero por claridad)
    ws.Visible = xlSheetVisible
    
    lineaError = 350
    
    ' Configuración adicional para compatibilidad con entornos cloud
    ' Asegurar que la hoja no esté protegida
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    ' Retornar referencia a la hoja creada
    Set fun802_CrearHojaDelimitadores = ws
    
    lineaError = 360
    
    Exit Function
    
ErrorHandler:
    Set fun802_CrearHojaDelimitadores = Nothing
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun802_CrearHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "PARÁMETRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 803: HACER HOJA VISIBLE
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Verifica la visibilidad de una hoja y la hace visible si está oculta
' Parámetros: ws (Worksheet)
' Retorna: Nada (Sub procedure)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Sub fun803_HacerHojaVisible(ws As Worksheet)
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 400
    
    ' Verificar parámetro de entrada
    If ws Is Nothing Then
        Exit Sub
    End If
    
    lineaError = 410
    
    ' Verificar que el libro permite cambiar visibilidad (no protegido)
    If ws.Parent.ProtectStructure Then
        ' Si el libro está protegido, no podemos cambiar la visibilidad
        ' Salir sin error porque la hoja podría estar ya visible
        Exit Sub
    End If
    
    lineaError = 420
    
    ' Verificar el estado actual de visibilidad y actuar según corresponda
    Select Case ws.Visible
        Case xlSheetVisible
            ' La hoja ya está visible, no hacer nada
            
        Case xlSheetHidden, xlSheetVeryHidden
            ' La hoja está oculta, hacerla visible
            ws.Visible = xlSheetVisible
            
        Case Else
            ' Estado desconocido, forzar visibilidad
            ws.Visible = xlSheetVisible
    End Select
    
    lineaError = 430
    
    Exit Sub
    
ErrorHandler:
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun803_HacerHojaVisible" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub

' =============================================================================
' FUNCIÓN AUXILIAR 804: LIMPIAR CONTENIDO DE HOJA
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Limpia todo el contenido de una hoja específica
' Parámetros: ws (Worksheet)
' Retorna: Nada (Sub procedure)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Sub fun804_LimpiarContenidoHoja(ws As Worksheet)
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 500
    
    ' Verificar parámetro de entrada
    If ws Is Nothing Then
        Exit Sub
    End If
    
    lineaError = 510
    
    ' Verificar que la hoja no esté protegida
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    lineaError = 520
    
    ' Limpiar todo el contenido de la hoja (método compatible con todas las versiones)
    ws.Cells.Clear
    
    lineaError = 530
    
    Exit Sub
    
ErrorHandler:
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun804_LimpiarContenidoHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub

' =============================================================================
' FUNCIÓN AUXILIAR 805: DETECTAR USE SYSTEM SEPARATORS
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Detecta si Excel está usando separadores del sistema
' Parámetros: Ninguno
' Retorna: String ("True" o "False")
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun805_DetectarUseSystemSeparators() As String
    
    On Error GoTo ErrorHandler
    
    ' Variable para almacenar el resultado
    Dim resultado As String
    Dim lineaError As Long
    
    lineaError = 600
    
    ' Detectar configuración actual de Use System Separators
    ' Usar compilación condicional para compatibilidad con versiones
    
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
    ' En caso de error, intentar método alternativo
    fun805_DetectarUseSystemSeparators = fun809_DetectarUseSystemSeparatorsLegacy()
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun805_DetectarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 806: DETECTAR DECIMAL SEPARATOR
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Detecta el separador decimal actual de Excel
' Parámetros: Ninguno
' Retorna: String (carácter del separador decimal)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun806_DetectarDecimalSeparator() As String
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    
    ' Detectar separador decimal actual (compatible con todas las versiones)
    fun806_DetectarDecimalSeparator = Application.DecimalSeparator
    
    lineaError = 710
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar método alternativo
    fun806_DetectarDecimalSeparator = fun810_DetectarDecimalSeparatorLegacy()
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun806_DetectarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 807: DETECTAR THOUSANDS SEPARATOR
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Detecta el separador de miles actual de Excel
' Parámetros: Ninguno
' Retorna: String (carácter del separador de miles)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun807_DetectarThousandsSeparator() As String
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    
    ' Detectar separador de miles actual (compatible con todas las versiones)
    fun807_DetectarThousandsSeparator = Application.ThousandsSeparator
    
    lineaError = 810
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar método alternativo
    fun807_DetectarThousandsSeparator = fun811_DetectarThousandsSeparatorLegacy()
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun807_DetectarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 808: OCULTAR HOJA DE DELIMITADORES
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Oculta la hoja de delimitadores si está habilitada la opción
' Parámetros: ws (Worksheet)
' Retorna: Nada (Sub procedure)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Sub fun808_OcultarHojaDelimitadores(ws As Worksheet)
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 900
    
    ' Verificar parámetro de entrada
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
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun808_OcultarHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub

' =============================================================================
' FUNCIÓN AUXILIAR 809: DETECTAR USE SYSTEM SEPARATORS (MÉTODO LEGACY)
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Método alternativo para detectar Use System Separators en versiones antiguas
' Parámetros: Ninguno
' Retorna: String ("True" o "False")
' Compatibilidad: Excel 97, 2003
' =============================================================================

Public Function fun809_DetectarUseSystemSeparatorsLegacy() As String
    
    On Error GoTo ErrorHandler
    
    ' Variables para comparación
    Dim separadorSistema As String
    Dim separadorExcel As String
    Dim lineaError As Long
    
    lineaError = 1000
    
    ' Obtener separador decimal del sistema (Windows)
    ' Método compatible con Excel 97 y 2003
    separadorSistema = Mid(CStr(1.1), 2, 1)
    
    lineaError = 1010
    
    ' Obtener separador decimal de Excel
    separadorExcel = Application.DecimalSeparator
    
    lineaError = 1020
    
    ' Si coinciden, probablemente Use System Separators está activado
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
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun809_DetectarUseSystemSeparatorsLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 810: DETECTAR DECIMAL SEPARATOR (MÉTODO LEGACY)
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Método alternativo para detectar separador decimal en versiones antiguas
' Parámetros: Ninguno
' Retorna: String (carácter del separador decimal)
' Compatibilidad: Excel 97, 2003
' =============================================================================

Public Function fun810_DetectarDecimalSeparatorLegacy() As String
    
    On Error GoTo ErrorHandler
    
    ' Variables para detección
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1100
    
    ' Método alternativo: formatear un número y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = CStr(1.1)
    
    lineaError = 1110
    
    ' El separador decimal es el segundo carácter en el formato estándar
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
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun810_DetectarDecimalSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 811: DETECTAR THOUSANDS SEPARATOR (MÉTODO LEGACY)
' =============================================================================
' Fecha: 2025-05-26 17:43:59 UTC
' Descripción: Método alternativo para detectar separador de miles en versiones antiguas
' Parámetros: Ninguno
' Retorna: String (carácter del separador de miles)
' Compatibilidad: Excel 97, 2003
' =============================================================================

Public Function fun811_DetectarThousandsSeparatorLegacy() As String
    
    On Error GoTo ErrorHandler
    
    ' Variables para detección
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1200
    
    ' Método alternativo: formatear un número grande y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = Format(1000, "#,##0")
    
    lineaError = 1210
    
    ' El separador de miles es el segundo carácter en números de 4 dígitos
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
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun811_DetectarThousandsSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

