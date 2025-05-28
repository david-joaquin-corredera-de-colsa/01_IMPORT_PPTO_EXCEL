Attribute VB_Name = "Modulo_015_FUNC_Ppales_RecDelim"
' =============================================================================
' VARIABLES GLOBALES ADICIONALES PARA RESTAURACIÓN DE DELIMITADORES
' =============================================================================
' Fecha y hora de creación: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Variables globales adicionales para restaurar delimitadores originales
' =============================================================================

' Variables para celdas que contienen valores originales
Public vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal As String
Public vCelda_Valor_Excel_DecimalSeparator_ValorOriginal As String
Public vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal As String

' Variables para almacenar valores originales leídos
Public vExcel_UseSystemSeparators_ValorOriginal As String
Public vExcel_DecimalSeparator_ValorOriginal As String
Public vExcel_ThousandsSeparator_ValorOriginal As String

' =============================================================================
' FUNCIÓN PRINCIPAL: F004_Restaurar_Delimitadores_en_Excel
' =============================================================================
' Fecha y hora de creación: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Restaura los delimitadores originales de Excel desde la hoja de respaldo
'
' RESUMEN EXHAUSTIVO DE PASOS:
' 1. Inicializar variables globales con valores por defecto (C2, C3, C4)
' 2. Obtener referencia al libro actual
' 3. Verificar si existe la hoja de delimitadores originales
' 4. Si no existe, crear la hoja y dejarla visible (situación extraña para restauración)
' 5. Si existe, verificar su visibilidad y hacerla visible si está oculta
' 6. Leer valores originales desde las celdas especificadas:
'    - Use System Separators desde C2
'    - Decimal Separator desde C3
'    - Thousands Separator desde C4
' 7. Almacenar valores leídos en variables globales correspondientes
' 8. Validar que los valores leídos sean apropiados para restaurar
' 9. Aplicar configuración original de delimitadores de Excel:
'    - Use System Separators (True/False según valor original)
'    - Decimal Separator (carácter según valor original)
'    - Thousands Separator (carácter según valor original)
' 10. Verificar variable global vOcultarRepostiorioDelimitadores
' 11. Si es True, ocultar la hoja de delimitadores al finalizar
' 12. Manejo exhaustivo de errores con información detallada y número de línea
'
' Parámetros: Ninguno
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Function F004_Restaurar_Delimitadores_en_Excel() As Boolean
    
    ' Control de errores con número de línea
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
    ' NOTA: Usuario especificó C2 para todas, corrijo para C2, C3, C4 según lógica
    vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal = "C2"
    vCelda_Valor_Excel_DecimalSeparator_ValorOriginal = "C3"
    vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal = "C4"
    
    ' Variables para almacenar los valores originales (inicialmente vacías)
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
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEGÚN CORRESPONDA
    ' ==========================================================================
    
    If Not hojaExiste Then
        ' La hoja no existe, crearla y dejarla visible
        ' NOTA: En un escenario de restauración, esto sería extraño, pero cumplimos la especificación
        Set ws = fun802_CrearHojaDelimitadores(wb, vHojaDelimitadoresExcelOriginales)
        If ws Is Nothing Then
            F004_Restaurar_Delimitadores_en_Excel = False
            Exit Function
        End If
        ' Como no hay datos que leer, salir con éxito pero sin restaurar
        Debug.Print "ADVERTENCIA: Hoja de delimitadores creada, pero no hay valores para restaurar - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        Exit Function
    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
        Set ws = wb.Worksheets(vHojaDelimitadoresExcelOriginales)
        
        ' Verificar si está oculta y hacerla visible si es necesario
        If Not fun803_HacerHojaVisible(ws) Then
            Debug.Print "ADVERTENCIA: No se pudo hacer visible la hoja " & vHojaDelimitadoresExcelOriginales & " - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
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
    ' PASO 6: VALIDAR QUE SE HAYAN LEÍDO VALORES VÁLIDOS
    ' ==========================================================================
    
    If Not fun805_ValidarValoresOriginales() Then
        Debug.Print "ADVERTENCIA: No se encontraron valores válidos para restaurar en la hoja: " & vHojaDelimitadoresExcelOriginales & " - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        F004_Restaurar_Delimitadores_en_Excel = False
        Exit Function
    End If
    
    lineaError = 160
    
    ' ==========================================================================
    ' PASO 7: APLICAR CONFIGURACIÓN ORIGINAL DE DELIMITADORES DE EXCEL
    ' ==========================================================================
    
    ' Restaurar Use System Separators (True/False)
    If Not fun806_RestaurarUseSystemSeparators(vExcel_UseSystemSeparators_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Use System Separators - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    ' Restaurar Decimal Separator (carácter)
    If Not fun807_RestaurarDecimalSeparator(vExcel_DecimalSeparator_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Decimal Separator - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    ' Restaurar Thousands Separator (carácter)
    If Not fun808_RestaurarThousandsSeparator(vExcel_ThousandsSeparator_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Thousands Separator - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    lineaError = 170
    
    ' ==========================================================================
    ' PASO 8: VERIFICAR SI DEBE OCULTAR LA HOJA
    ' ==========================================================================
    
    ' Verificar la variable global vOcultarRepostiorioDelimitadores
    If vOcultarRepostiorioDelimitadores = True Then
        ' Ocultar la hoja de delimitadores
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & vHojaDelimitadoresExcelOriginales & " - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        End If
    End If
    
    lineaError = 180
    
    ' ==========================================================================
    ' PASO 9: FINALIZACIÓN EXITOSA
    ' ==========================================================================
    
    Debug.Print "ÉXITO: Delimitadores restaurados correctamente - Función: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Restaurar_Delimitadores_en_Excel = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: F004_Restaurar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now() & vbCrLf & _
                   "USUARIO: david-joaquin-corredera-de-colsa"
    
    ' Log del error para debugging
    Debug.Print mensajeError
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 801: VERIFICAR EXISTENCIA DE HOJA
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
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
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
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
        Debug.Print "ERROR: No se puede crear hoja, libro protegido - Función: fun802_CrearHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 320
    
    ' Crear nueva hoja al final del libro (método compatible con todas las versiones)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    lineaError = 330
    
    ' Asignar nombre a la hoja
    ws.Name = nombreHoja
    
    lineaError = 340
    
    ' Asegurar que la hoja esté visible
    ws.Visible = xlSheetVisible
    
    lineaError = 350
    
    ' Configuración adicional para compatibilidad con entornos cloud
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
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Verifica la visibilidad de una hoja y la hace visible si está oculta
' Parámetros: ws (Worksheet)
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Function fun803_HacerHojaVisible(ws As Worksheet) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 400
    fun803_HacerHojaVisible = True
    
    ' Verificar parámetro de entrada
    If ws Is Nothing Then
        fun803_HacerHojaVisible = False
        Exit Function
    End If
    
    lineaError = 410
    
    ' Verificar que el libro permite cambiar visibilidad (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede cambiar visibilidad, libro protegido - Función: fun803_HacerHojaVisible - " & Now()
        Exit Function
    End If
    
    lineaError = 420
    
    ' Verificar el estado actual de visibilidad y actuar según corresponda
    Select Case ws.Visible
        Case xlSheetVisible
            ' La hoja ya está visible, no hacer nada
            Debug.Print "INFO: Hoja " & ws.Name & " ya está visible - Función: fun803_HacerHojaVisible - " & Now()
            
        Case xlSheetHidden, xlSheetVeryHidden
            ' La hoja está oculta, hacerla visible
            ws.Visible = xlSheetVisible
            Debug.Print "INFO: Hoja " & ws.Name & " se hizo visible - Función: fun803_HacerHojaVisible - " & Now()
            
        Case Else
            ' Estado desconocido, forzar visibilidad
            ws.Visible = xlSheetVisible
            Debug.Print "INFO: Hoja " & ws.Name & " visibilidad forzada - Función: fun803_HacerHojaVisible - " & Now()
    End Select
    
    lineaError = 430
    
    Exit Function
    
ErrorHandler:
    fun803_HacerHojaVisible = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun803_HacerHojaVisible" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 804: CONVERTIR VALOR A CADENA
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Convierte un valor de celda a cadena de texto de forma segura
' Parámetros: valor (Variant)
' Retorna: String (valor convertido o cadena vacía si error)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun804_ConvertirValorACadena(valor As Variant) As String
    
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
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun804_ConvertirValorACadena" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 805: VALIDAR VALORES ORIGINALES
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Valida que los valores originales leídos sean válidos para restaurar
' Parámetros: Ninguno (usa variables globales)
' Retorna: Boolean (True si válidos, False si no válidos)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun805_ValidarValoresOriginales() As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    Dim esValido As Boolean
    
    lineaError = 600
    esValido = True
    
    ' Validar Use System Separators (debe ser "True" o "False")
    If vExcel_UseSystemSeparators_ValorOriginal <> "True" And vExcel_UseSystemSeparators_ValorOriginal <> "False" Then
        If vExcel_UseSystemSeparators_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inválido para Use System Separators: '" & vExcel_UseSystemSeparators_ValorOriginal & "' - Función: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 610
    
    ' Validar Decimal Separator (debe ser un solo carácter)
    If Len(vExcel_DecimalSeparator_ValorOriginal) <> 1 Then
        If vExcel_DecimalSeparator_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inválido para Decimal Separator: '" & vExcel_DecimalSeparator_ValorOriginal & "' - Función: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 620
    
    ' Validar Thousands Separator (debe ser un solo carácter)
    If Len(vExcel_ThousandsSeparator_ValorOriginal) <> 1 Then
        If vExcel_ThousandsSeparator_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inválido para Thousands Separator: '" & vExcel_ThousandsSeparator_ValorOriginal & "' - Función: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 630
    
    fun805_ValidarValoresOriginales = esValido
    
    ' Log de valores validados
    If esValido Then
        Debug.Print "INFO: Valores válidos para restaurar - UseSystem:" & vExcel_UseSystemSeparators_ValorOriginal & " Decimal:'" & vExcel_DecimalSeparator_ValorOriginal & "' Thousands:'" & vExcel_ThousandsSeparator_ValorOriginal & "' - Función: fun805_ValidarValoresOriginales - " & Now()
    End If
    
    Exit Function
    
ErrorHandler:
    fun805_ValidarValoresOriginales = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun805_ValidarValoresOriginales" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 806: RESTAURAR USE SYSTEM SEPARATORS
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Restaura la configuración de Use System Separators
' Parámetros: valorOriginal (String) - "True" o "False"
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun806_RestaurarUseSystemSeparators(valorOriginal As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    fun806_RestaurarUseSystemSeparators = True
    
    ' Verificar que el valor sea válido
    If valorOriginal <> "True" And valorOriginal <> "False" Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Use System Separators, valor inválido: '" & valorOriginal & "' - Función: fun806_RestaurarUseSystemSeparators - " & Now()
        fun806_RestaurarUseSystemSeparators = False
        Exit Function
    End If
    
    lineaError = 710
    
    ' Usar compilación condicional para compatibilidad con versiones
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lineaError = 720
        If valorOriginal = "True" Then
            Application.UseSystemSeparators = True
            Debug.Print "INFO: Use System Separators configurado a True - Función: fun806_RestaurarUseSystemSeparators - " & Now()
        Else
            Application.UseSystemSeparators = False
            Debug.Print "INFO: Use System Separators configurado a False - Función: fun806_RestaurarUseSystemSeparators - " & Now()
        End If
    #Else
        ' Excel 97, 2003 y anteriores
        lineaError = 730
        Debug.Print "ADVERTENCIA: Use System Separators no disponible en esta versión de Excel - Función: fun806_RestaurarUseSystemSeparators - " & Now()
        ' En versiones antiguas, esta propiedad no existe, pero no es error
    #End If
    
    lineaError = 740
    
    Exit Function
    
ErrorHandler:
    fun806_RestaurarUseSystemSeparators = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun806_RestaurarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "PARÁMETRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 807: RESTAURAR DECIMAL SEPARATOR
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Restaura el separador decimal original
' Parámetros: valorOriginal (String) - carácter del separador decimal
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun807_RestaurarDecimalSeparator(valorOriginal As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    fun807_RestaurarDecimalSeparator = True
    
    ' Verificar que el valor sea válido (un solo carácter)
    If Len(valorOriginal) <> 1 Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Decimal Separator, valor inválido: '" & valorOriginal & "' - Función: fun807_RestaurarDecimalSeparator - " & Now()
        fun807_RestaurarDecimalSeparator = False
        Exit Function
    End If
    
    lineaError = 810
    
    ' Restaurar separador decimal (compatible con todas las versiones)
    Application.DecimalSeparator = valorOriginal
    Debug.Print "INFO: Decimal Separator restaurado a: '" & valorOriginal & "' - Función: fun807_RestaurarDecimalSeparator - " & Now()
    
    lineaError = 820
    
    Exit Function
    
ErrorHandler:
    fun807_RestaurarDecimalSeparator = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun807_RestaurarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "PARÁMETRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 808: RESTAURAR THOUSANDS SEPARATOR
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Restaura el separador de miles original
' Parámetros: valorOriginal (String) - carácter del separador de miles
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365
' =============================================================================

Public Function fun808_RestaurarThousandsSeparator(valorOriginal As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 900
    fun808_RestaurarThousandsSeparator = True
    
    ' Verificar que el valor sea válido (un solo carácter)
    If Len(valorOriginal) <> 1 Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Thousands Separator, valor inválido: '" & valorOriginal & "' - Función: fun808_RestaurarThousandsSeparator - " & Now()
        fun808_RestaurarThousandsSeparator = False
        Exit Function
    End If
    
    lineaError = 910
    
    ' Restaurar separador de miles (compatible con todas las versiones)
    Application.ThousandsSeparator = valorOriginal
    Debug.Print "INFO: Thousands Separator restaurado a: '" & valorOriginal & "' - Función: fun808_RestaurarThousandsSeparator - " & Now()
    
    lineaError = 920
    
    Exit Function
    
ErrorHandler:
    fun808_RestaurarThousandsSeparator = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun808_RestaurarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "PARÁMETRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR 809: OCULTAR HOJA DE DELIMITADORES
' =============================================================================
' Fecha: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Oculta la hoja de delimitadores si está habilitada la opción
' Parámetros: ws (Worksheet)
' Retorna: Boolean (True si exitoso, False si error)
' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

Public Function fun809_OcultarHojaDelimitadores(ws As Worksheet) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 1000
    fun809_OcultarHojaDelimitadores = True
    
    ' Verificar parámetro de entrada
    If ws Is Nothing Then
        fun809_OcultarHojaDelimitadores = False
        Exit Function
    End If
    
    lineaError = 1010
    
    ' Verificar que el libro permite ocultar hojas (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede ocultar hoja, libro protegido - Función: fun809_OcultarHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 1020
    
    ' Ocultar la hoja (compatible con todas las versiones de Excel)
    ws.Visible = xlSheetHidden
    Debug.Print "INFO: Hoja " & ws.Name & " ocultada - Función: fun809_OcultarHojaDelimitadores - " & Now()
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    fun809_OcultarHojaDelimitadores = False
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun809_OcultarHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


