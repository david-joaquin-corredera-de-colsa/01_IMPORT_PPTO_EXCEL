Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_08"
Option Explicit

Public Function fun807_RestaurarDecimalSeparator(ByVal valorOriginal As String) As Boolean
    
    ' =============================================================================
    ' FUNCIÓN AUXILIAR 807: RESTAURAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha y hora de creación: 2025-06-16 22:33:04 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripción: Restaura la configuración del separador decimal
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim valorActualAntes As String
    Dim valorActualDespues As String
    Dim caracterASCII As Integer
    
    ' Inicialización
    strFuncion = "fun807_RestaurarDecimalSeparator"
    fun807_RestaurarDecimalSeparator = False
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y CAPTURA DE ESTADO ACTUAL
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando restauración de DecimalSeparator"
    valorActualAntes = Application.DecimalSeparator
    caracterASCII = Asc(valorOriginal)
    
    fun801_LogMessage "[INICIO] " & strFuncion & " - Valor a restaurar: '" & valorOriginal & "' | ASCII: " & caracterASCII & _
                      " | Valor actual: '" & valorActualAntes & "' (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 2: VALIDAR PARÁMETRO DE ENTRADA
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Validando parámetro de entrada"
    fun801_LogMessage "[PASO 1] " & strContexto & " - Verificando formato y longitud (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) = 0 Then
        strMensajeError = "El valor original para DecimalSeparator no puede estar vacío"
        Err.Raise ERROR_BASE_IMPORT + 1016, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[DETALLE] Longitud del valor: " & Len(valorOriginal) & " | Valor trimmed: '" & Trim(valorOriginal) & "'", False, "", strFuncion
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: VERIFICAR QUE SEA UN CARÁCTER VÁLIDO
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Verificando validez del carácter"
    fun801_LogMessage "[PASO 2] " & strContexto & " - Validando formato de separador decimal (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) <> 1 Then
        strMensajeError = "DecimalSeparator debe ser exactamente un carácter - Longitud recibida: " & Len(valorOriginal) & " | Valor: '" & valorOriginal & "'"
        Err.Raise ERROR_BASE_IMPORT + 1017, strFuncion, strMensajeError
    End If
    
    If Not (valorOriginal = "." Or valorOriginal = ",") Then
        strMensajeError = "DecimalSeparator no es válido - Esperado: '.' o ',' | Recibido: '" & valorOriginal & "' | ASCII: " & caracterASCII
        Err.Raise ERROR_BASE_IMPORT + 1018, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[VALIDACIÓN] Carácter válido confirmado: '" & valorOriginal & "'", False, "", strFuncion
    
    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: APLICAR CONFIGURACIÓN A EXCEL
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Aplicando nuevo separador decimal a Excel"
    fun801_LogMessage "[PASO 3] " & strContexto & " - Cambiando de '" & valorActualAntes & "' a '" & valorOriginal & "' (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    Application.DecimalSeparator = valorOriginal
    fun801_LogMessage "[APLICADO] Application.DecimalSeparator configurado", False, "", strFuncion
    
    lngLineaError = 160
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR QUE EL CAMBIO SE APLICÓ CORRECTAMENTE
    '--------------------------------------------------------------------------
    lngLineaError = 170
    strContexto = "Verificando aplicación correcta del separador"
    fun801_LogMessage "[PASO 4] " & strContexto & " (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    valorActualDespues = Application.DecimalSeparator
    fun801_LogMessage "[VERIFICACIÓN] Separador después del cambio: '" & valorActualDespues & "' | Esperado: '" & valorOriginal & "'", False, "", strFuncion
    
    If valorActualDespues = valorOriginal Then
        fun801_LogMessage "[ÉXITO] DecimalSeparator restaurado exitosamente", False, "", strFuncion
        fun807_RestaurarDecimalSeparator = True
    Else
        strMensajeError = "Error en verificación de DecimalSeparator - Valor esperado: '" & valorOriginal & "' | Valor actual: '" & valorActualDespues & "'"
        Err.Raise ERROR_BASE_IMPORT + 1019, strFuncion, strMensajeError
    End If
    
    lngLineaError = 180
    
    ' Verificación adicional con formato de número
    Dim numeroTest As Double
    Dim numeroFormateado As String
    numeroTest = 1.23
    numeroFormateado = CStr(numeroTest)
    fun801_LogMessage "[VERIFICACIÓN ADICIONAL] Número test: " & numeroTest & " | Formateado: '" & numeroFormateado & "' | Separador detectado: '" & Mid(numeroFormateado, 2, 1) & "'", False, "", strFuncion
    
    fun801_LogMessage "[FINALIZACIÓN] " & strFuncion & " completado exitosamente - Cambio aplicado: '" & valorActualAntes & "' ? '" & valorOriginal & "'", False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "Línea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "Número de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripción VBA: " & Err.Description & vbCrLf & _
                      "Valor Original: '" & valorOriginal & "'" & vbCrLf & _
                      "ASCII del valor: " & caracterASCII & vbCrLf & _
                      "Valor Antes: '" & valorActualAntes & "'" & vbCrLf & _
                      "Valor Después: '" & valorActualDespues & "'" & vbCrLf & _
                      "Longitud valor: " & Len(valorOriginal) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun807_RestaurarDecimalSeparator = False
End Function

Public Function fun808_RestaurarThousandsSeparator(ByVal valorOriginal As String) As Boolean
    
    ' =============================================================================
    ' FUNCIÓN AUXILIAR 808: RESTAURAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha y hora de creación: 2025-06-16 22:33:04 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripción: Restaura la configuración del separador de miles
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim valorActualAntes As String
    Dim valorActualDespues As String
    Dim caracterASCII As Integer
    Dim esCaracterValido As Boolean
    
    ' Inicialización
    strFuncion = "fun808_RestaurarThousandsSeparator"
    fun808_RestaurarThousandsSeparator = False
    lngLineaError = 0
    strContexto = ""
    esCaracterValido = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y CAPTURA DE ESTADO ACTUAL
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando restauración de ThousandsSeparator"
    valorActualAntes = Application.ThousandsSeparator
    
    If Len(valorOriginal) > 0 Then
        caracterASCII = Asc(valorOriginal)
    Else
        caracterASCII = 0
    End If
    
    fun801_LogMessage "[INICIO] " & strFuncion & " - Valor a restaurar: '" & valorOriginal & "' | ASCII: " & caracterASCII & _
                      " | Valor actual: '" & valorActualAntes & "' (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 2: VALIDAR PARÁMETRO DE ENTRADA
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Validando parámetro de entrada"
    fun801_LogMessage "[PASO 1] " & strContexto & " - Verificando formato y longitud (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) = 0 Then
        strMensajeError = "El valor original para ThousandsSeparator no puede estar vacío"
        Err.Raise ERROR_BASE_IMPORT + 1020, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[DETALLE] Longitud del valor: " & Len(valorOriginal) & " | Valor trimmed: '" & Trim(valorOriginal) & "' | ASCII: " & caracterASCII, False, "", strFuncion
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: VERIFICAR QUE SEA UN CARÁCTER VÁLIDO
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Verificando validez del carácter separador"
    fun801_LogMessage "[PASO 2] " & strContexto & " - Validando formato de separador de miles (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) <> 1 Then
        strMensajeError = "ThousandsSeparator debe ser exactamente un carácter - Longitud recibida: " & Len(valorOriginal) & " | Valor: '" & valorOriginal & "'"
        Err.Raise ERROR_BASE_IMPORT + 1021, strFuncion, strMensajeError
    End If
    
    ' Verificar caracteres válidos para separador de miles
    If valorOriginal = "." Then
        esCaracterValido = True
        fun801_LogMessage "[VALIDACIÓN] Carácter punto (.) detectado", False, "", strFuncion
    ElseIf valorOriginal = "," Then
        esCaracterValido = True
        fun801_LogMessage "[VALIDACIÓN] Carácter coma (,) detectado", False, "", strFuncion
    ElseIf valorOriginal = " " Then
        esCaracterValido = True
        fun801_LogMessage "[VALIDACIÓN] Carácter espacio ( ) detectado | ASCII: 32", False, "", strFuncion
    ElseIf valorOriginal = Chr(39) Then ' Comilla simple
        esCaracterValido = True
        fun801_LogMessage "[VALIDACIÓN] Carácter comilla simple (') detectado | ASCII: 39", False, "", strFuncion
    Else
        esCaracterValido = False
        fun801_LogMessage "[VALIDACIÓN] Carácter no reconocido: '" & valorOriginal & "' | ASCII: " & caracterASCII, False, "", strFuncion
    End If
    
    If Not esCaracterValido Then
        strMensajeError = "ThousandsSeparator no es válido - Esperado: '.', ',', ' ' o ''' | Recibido: '" & valorOriginal & "' | ASCII: " & caracterASCII
        Err.Raise ERROR_BASE_IMPORT + 1022, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[VALIDACIÓN] Carácter válido confirmado: '" & valorOriginal & "'", False, "", strFuncion
    
    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: APLICAR CONFIGURACIÓN A EXCEL
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Aplicando nuevo separador de miles a Excel"
    fun801_LogMessage "[PASO 3] " & strContexto & " - Cambiando de '" & valorActualAntes & "' a '" & valorOriginal & "' (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    Application.ThousandsSeparator = valorOriginal
    fun801_LogMessage "[APLICADO] Application.ThousandsSeparator configurado", False, "", strFuncion
    
    lngLineaError = 160
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR QUE EL CAMBIO SE APLICÓ CORRECTAMENTE
    '--------------------------------------------------------------------------
    lngLineaError = 170
    strContexto = "Verificando aplicación correcta del separador"
    fun801_LogMessage "[PASO 4] " & strContexto & " (Línea: " & lngLineaError & ")", False, "", strFuncion
    
    valorActualDespues = Application.ThousandsSeparator
    fun801_LogMessage "[VERIFICACIÓN] Separador después del cambio: '" & valorActualDespues & "' | ASCII: " & Asc(valorActualDespues) & " | Esperado: '" & valorOriginal & "'", False, "", strFuncion
    
    If valorActualDespues = valorOriginal Then
        fun801_LogMessage "[ÉXITO] ThousandsSeparator restaurado exitosamente", False, "", strFuncion
        fun808_RestaurarThousandsSeparator = True
    Else
        strMensajeError = "Error en verificación de ThousandsSeparator - Valor esperado: '" & valorOriginal & "' (ASCII:" & caracterASCII & ") | Valor actual: '" & valorActualDespues & "' (ASCII:" & Asc(valorActualDespues) & ")"
        Err.Raise ERROR_BASE_IMPORT + 1023, strFuncion, strMensajeError
    End If
    
    lngLineaError = 180
    
    ' Verificación adicional con formato de número
    Dim numeroTest As Long
    Dim numeroFormateado As String
    numeroTest = 12345
    numeroFormateado = Format(numeroTest, "#,##0")
    fun801_LogMessage "[VERIFICACIÓN ADICIONAL] Número test: " & numeroTest & " | Formateado: '" & numeroFormateado & "' | Longitud: " & Len(numeroFormateado), False, "", strFuncion
    
    If Len(numeroFormateado) >= 6 Then ' 12,345 = 6 caracteres mínimo
        fun801_LogMessage "[VERIFICACIÓN ADICIONAL] Separador en formato: '" & Mid(numeroFormateado, 3, 1) & "' | ASCII: " & Asc(Mid(numeroFormateado, 3, 1)), False, "", strFuncion
    End If
    
    fun801_LogMessage "[FINALIZACIÓN] " & strFuncion & " completado exitosamente - Cambio aplicado: '" & valorActualAntes & "' ? '" & valorOriginal & "'", False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "Línea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "Número de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripción VBA: " & Err.Description & vbCrLf & _
                      "Valor Original: '" & valorOriginal & "'" & vbCrLf & _
                      "ASCII del valor: " & caracterASCII & vbCrLf & _
                      "Carácter válido: " & esCaracterValido & vbCrLf & _
                      "Valor Antes: '" & valorActualAntes & "'" & vbCrLf & _
                      "Valor Después: '" & valorActualDespues & "'" & vbCrLf & _
                      "Longitud valor: " & Len(valorOriginal) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun808_RestaurarThousandsSeparator = False
End Function


