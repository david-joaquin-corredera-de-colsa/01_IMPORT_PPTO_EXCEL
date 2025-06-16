Attribute VB_Name = "Modulo_012_FUNC_Principales_02"
Option Explicit


Public Function F005_Procesar_Hoja_Comprobacion() As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F005_Procesar_Hoja_Comprobacion
    ' Fecha y Hora de Creación: 2025-06-01 21:52:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Copia todo el contenido de la hoja de envío a la hoja de comprobación
    ' para permitir verificación y control de calidad de los datos procesados.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que las hojas de envío y comprobación existan
    ' 2. Obtener referencias a las hojas de trabajo
    ' 3. Copiar contenido completo de hoja envío a hoja comprobación
    ' 4. Verificar que la copia se realizó correctamente
    ' 5. Registrar el resultado en el log del sistema
    '
    ' Parámetros: Ninguno (usa variables globales)
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas de trabajo
    Dim wsEnvio As Worksheet
    Dim wsComprobacion As Worksheet
    
    ' Inicialización
    strFuncion = "F005_Procesar_Hoja_Comprobacion" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F005_Procesar_Hoja_Comprobacion"
    F005_Procesar_Hoja_Comprobacion = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que las hojas existan
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando existencia de hojas para procesamiento de comprobación...", False, "", strFuncion
    
    ' Validar hoja de envío
    If Not fun802_SheetExists(gstrNuevaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 301, strFuncion, _
            "La hoja de envío no existe: " & gstrNuevaHojaImportacion_Envio
    End If
    
    ' Validar hoja de comprobación
    If Not fun802_SheetExists(gstrNuevaHojaImportacion_Comprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 302, strFuncion, _
            "La hoja de comprobación no existe: " & gstrNuevaHojaImportacion_Comprobacion
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Obtener referencias a las hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Obteniendo referencias a hojas de trabajo...", False, "", strFuncion
    
    Set wsEnvio = ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio)
    Set wsComprobacion = ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Comprobacion)
    
    ' Verificar que las referencias son válidas
    If wsEnvio Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 303, strFuncion, _
            "No se pudo obtener referencia a la hoja de envío"
    End If
    
    If wsComprobacion Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 304, strFuncion, _
            "No se pudo obtener referencia a la hoja de comprobación"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Copiar contenido completo de hoja envío a hoja comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Copiando contenido de hoja de envío a hoja de comprobación...", _
                      False, gstrNuevaHojaImportacion_Envio, gstrNuevaHojaImportacion_Comprobacion
    
    If Not fun817_CopiarContenidoCompleto(wsEnvio, wsComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 305, strFuncion, _
            "Error al copiar contenido de hoja envío a hoja comprobación"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Verificar que la copia se realizó correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Verificando integridad de la copia...", False, "", strFuncion
    
    ' Verificación básica: comparar si ambas hojas tienen contenido
    If wsEnvio.UsedRange Is Nothing And wsComprobacion.UsedRange Is Nothing Then
        ' Ambas están vacías, es correcto
        fun801_LogMessage "Verificación completada: ambas hojas están vacías (correcto)", False, "", strFuncion
    ElseIf wsEnvio.UsedRange Is Nothing Or wsComprobacion.UsedRange Is Nothing Then
        ' Una tiene contenido y la otra no, es un error
        Err.Raise ERROR_BASE_IMPORT + 306, strFuncion, _
            "Error en verificación: inconsistencia en contenido de hojas"
    Else
        ' Ambas tienen contenido, verificar que tienen el mismo rango
        If wsEnvio.UsedRange.Rows.Count = wsComprobacion.UsedRange.Rows.Count And _
           wsEnvio.UsedRange.Columns.Count = wsComprobacion.UsedRange.Columns.Count Then
            fun801_LogMessage "Verificación completada: dimensiones coinciden", False, "", strFuncion
        Else
            Err.Raise ERROR_BASE_IMPORT + 307, strFuncion, _
                "Error en verificación: las dimensiones de los rangos no coinciden"
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 90001
    fun801_LogMessage "Procesamiento de hoja de comprobación completado con éxito", _
                      False, gstrNuevaHojaImportacion_Envio, gstrNuevaHojaImportacion_Comprobacion
    
    F005_Procesar_Hoja_Comprobacion = True
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F005_Procesar_Hoja_Comprobacion = False
End Function




'Public Function F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio(ByVal strHojaComprobacion As String, ByVal strHojaEnvio As String) As Boolean
Public Function F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio(ByVal strHojaComprobacion As String, ByVal strHojaEnvio As String, _
    ByRef vScenario_xPL As String, ByRef vYear_xPL As String, ByRef vEntity_xPL As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio
    ' Fecha y Hora de Creación: 2025-06-03 00:14:44 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Copia datos específicos desde la hoja de comprobación hacia la hoja de envío,
    ' implementando lógica condicional basada en la comparación de rangos entre ambas hojas.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    ' 2. Detectar rangos de datos en hoja de comprobación
    ' 3. Detectar rangos de datos en hoja de envío
    ' 4. Comparar si los rangos son idénticos
    ' 5. Si rangos son iguales: copiar datos específicos (filas+2, columnas+11)
    ' 6. Si rangos son diferentes: copiar contenido completo y limpiar excesos
    ' 7. Verificar integridad de la operación
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Parámetros:
    ' - strHojaEnvio: Nombre de la hoja de destino (envío)
    ' - strHojaComprobacion: Nombre de la hoja de origen (comprobación)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    'Variable para habilitar/deshabilitar partes de esta funcion
    Dim vEnabled_Parts As Boolean
    'vEnabled_Parts = True
    'If vEnabled_Parts Then
    'End If 'vEnabled_Parts Then
    
    ' Variables para mostrar información de rangos
    Dim strMensajeRangosEnvio As String
    Dim strMensajeRangosComprobacion As String
    Dim strMensajeRangosCompleto As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas de trabajo
    Dim wsEnvio As Worksheet
    Dim wsComprobacion As Worksheet
    
    ' Variables para rangos de la hoja de comprobación
    Dim vFila_Inicial_HojaComprobacion As Long
    Dim vFila_Final_HojaComprobacion As Long
    Dim vColumna_Inicial_HojaComprobacion As Long
    Dim vColumna_Final_HojaComprobacion As Long
    
    ' Variables para rangos de la hoja de envío
    Dim vFila_Inicial_HojaEnvio As Long
    Dim vFila_Final_HojaEnvio As Long
    Dim vColumna_Inicial_HojaEnvio As Long
    Dim vColumna_Final_HojaEnvio As Long
    
    ' Variable para comparación de rangos
    Dim vLosRangosSonIguales As Boolean
    
    ' Variables para rangos de copia
    Dim rngOrigen As Range
    Dim rngDestino As Range
    
    ' Inicialización
    strFuncion = "F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio"
    F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio = False
    lngLineaError = 0
    vLosRangosSonIguales = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para comprobar datos enviados...", False, "", strFuncion
    
    ' Validar hoja de envío
    If Not fun802_SheetExists(strHojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de envío no existe: " & strHojaEnvio
    End If
    
    ' Validar hoja de comprobación
    If Not fun802_SheetExists(strHojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 702, strFuncion, _
            "La hoja de comprobación no existe: " & strHojaComprobacion
    End If
    
    ' Obtener referencias a las hojas
    Set wsEnvio = ThisWorkbook.Worksheets(strHojaEnvio)
    Set wsComprobacion = ThisWorkbook.Worksheets(strHojaComprobacion)
    
    ' Verificar que las referencias son válidas
    If wsEnvio Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja de envío"
    End If
    
    If wsComprobacion Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 704, strFuncion, _
            "No se pudo obtener referencia a la hoja de comprobación"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave específicas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento específico
    ' Solo si necesitas valores diferentes a los por defecto
    Dim vEscenarioAdmitido, vUltimoMesCarga As String
    vEscenarioAdmitido = UCase(Trim(CONST_ESCENARIO_ADMITIDO))
    vUltimoMesCarga = UCase(Trim(CONST_ULTIMO_MES_DE_CARGA))
    Call fun826_ConfigurarPalabrasClave(vEscenarioAdmitido, vEscenarioAdmitido, vEscenarioAdmitido, vUltimoMesCarga)
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Detectando rangos de datos en hoja de comprobación...", False, "", strHojaComprobacion
    
    If Not fun822_DetectarRangoCompletoHoja(wsComprobacion, _
                                           vFila_Inicial_HojaComprobacion, _
                                           vFila_Final_HojaComprobacion, _
                                           vColumna_Inicial_HojaComprobacion, _
                                           vColumna_Final_HojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 705, strFuncion, _
            "Error al detectar rangos en hoja de comprobación"
    End If
    
    fun801_LogMessage "Rangos de comprobación - Filas: " & vFila_Inicial_HojaComprobacion & " a " & vFila_Final_HojaComprobacion & _
                      ", Columnas: " & vColumna_Inicial_HojaComprobacion & " a " & vColumna_Final_HojaComprobacion, _
                      False, "", strHojaComprobacion
    
    vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaComprobacion - 1 'Le quitamos 1, para que considere también la fila en la que están los headers de los meses M01 ... M12
    vColumna_Final_HojaComprobacion = vColumna_Inicial_HojaComprobacion + 22
    
    '--------------------------------------------------------------------------
    ' 3. Detectar rangos de datos en hoja de envío
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Detectando rangos de datos en hoja de envío...", False, "", strHojaEnvio
    
    If Not fun822_DetectarRangoCompletoHoja(wsEnvio, _
                                           vFila_Inicial_HojaEnvio, _
                                           vFila_Final_HojaEnvio, _
                                           vColumna_Inicial_HojaEnvio, _
                                           vColumna_Final_HojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 706, strFuncion, _
            "Error al detectar rangos en hoja de envío"
    End If
    
    fun801_LogMessage "Rangos de envío - Filas: " & vFila_Inicial_HojaEnvio & " a " & vFila_Final_HojaEnvio & _
                      ", Columnas: " & vColumna_Inicial_HojaEnvio & " a " & vColumna_Final_HojaEnvio, _
                      False, "", strHojaEnvio
            
    vFila_Inicial_HojaEnvio = vFila_Inicial_HojaEnvio - 1 'Le quitamos 1, para que considere también la fila en la que están los headers de los meses M01 ... M12
    vColumna_Final_HojaEnvio = vColumna_Inicial_HojaEnvio + 22
            
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar información completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then

        lngLineaError = 125
        strMensajeRangosCompleto = "INFORMACIÓN COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE ENVÍO: " & strHojaEnvio & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaEnvio & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaEnvio & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaEnvio & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaEnvio & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaEnvio - vFila_Inicial_HojaEnvio + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaEnvio - vColumna_Inicial_HojaEnvio + 1) & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE COMPROBACIÓN: " & strHojaComprobacion & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaComprobacion & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaComprobacion & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaComprobacion & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaComprobacion & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaComprobacion - vFila_Inicial_HojaComprobacion + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaComprobacion - vColumna_Inicial_HojaComprobacion + 1)
        
        MsgBox strMensajeRangosCompleto, vbInformation, "Rangos Completos - " & strFuncion
        
    End If 'vEnabled_Parts Then
    
    '--------------------------------------------------------------------------
    ' 4. Comparar si los rangos son idénticos
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Comparando rangos entre hojas...", False, "", strFuncion
    
    If (vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaEnvio) And _
       (vFila_Final_HojaComprobacion = vFila_Final_HojaEnvio) And _
       (vColumna_Inicial_HojaComprobacion = vColumna_Inicial_HojaEnvio) And _
       (vColumna_Final_HojaComprobacion = vColumna_Final_HojaEnvio) Then
        vLosRangosSonIguales = True
        fun801_LogMessage "Los rangos son idénticos - Aplicando copia específica", False, "", strFuncion
    Else
        vLosRangosSonIguales = False
        fun801_LogMessage "Los rangos son diferentes - Aplicando copia completa", False, "", strFuncion
    End If
    
    'MsgBox "Los Rangos son Iguales? = " & vLosRangosSonIguales
    
    'En realidad si los rangos no salen iguales, tiene que ser
    '   porque en una de las 2 hojas esté considerando como "Contenido"
    '   algunas celdas que en realidad no tienen contenido
    '   (tendríamos que hacerle un ClearConents a algunos rangos,
    '   como por ejemplo columnas anteriores a la del primer "BUDGET_OS", columnas posteriores a la del "M12"
    '   o filas anteriores a la del M12
    
    'Asi que vamos a forzar a que los rangos sean iguales
    ' y vamos a usar los rangos de la strHojaComprobacion
    vLosRangosSonIguales = True
    
    '--------------------------------------------------------------------------
    ' 5. Procesar según el resultado de la comparación
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then
    '>>>>>
    'Deshabilitamos esta parte de la función
    '   porque para la comprobación de datos enviados
    '   esta parte de la función no tiene sentido
    
    If vLosRangosSonIguales = True Then
        '----------------------------------------------------------------------
        ' 5.1. Rangos iguales: Copiar datos específicos (filas+2, columnas+11)
        '----------------------------------------------------------------------
        lngLineaError = 90003
        fun801_LogMessage "Ejecutando copia específica para rangos idénticos...", False, "", strFuncion
        
        ' Validar que hay suficientes filas y columnas para el offset
        'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
           (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
            
            ' Definir rango origen (desde comprobación)
            Set rngOrigen = wsComprobacion.Range( _
                wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Definir rango destino (hacia envío)
            Set rngDestino = wsEnvio.Range( _
                wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Realizar copia de valores únicamente
            If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
                Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                    "Error al copiar valores específicos"
            End If
            
            fun801_LogMessage "Copia específica completada correctamente", False, "", strFuncion
        'Else
        '    fun801_LogMessage "Advertencia: Offset insuficiente para copia específica, omitiendo operación", False, "", strFuncion
        'End If
        
    Else
        '----------------------------------------------------------------------
        ' 5.2. Rangos diferentes: Copiar contenido completo de HojaComprobacion a HojaEnvio
        '----------------------------------------------------------------------
        lngLineaError = 100
        fun801_LogMessage "Ejecutando copia completa para rangos diferentes...", False, "", strFuncion
        
        ' Definir rango origen completo (desde comprobación)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino completo (hacia envío)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores únicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 708, strFuncion, _
                "Error al copiar contenido completo"
        End If
        
        '----------------------------------------------------------------------
        ' 5.3. Limpiar excesos en hoja de envío
        '----------------------------------------------------------------------
        lngLineaError = 110
        fun801_LogMessage "Limpiando excesos en hoja de envío...", False, "", strHojaEnvio
        
        ' Limpiar filas excedentes
        If Not fun824_LimpiarFilasExcedentes(wsEnvio, vFila_Final_HojaComprobacion) Then
            fun801_LogMessage "Advertencia: Error al limpiar filas excedentes", False, "", strHojaEnvio
        End If
        
        ' Limpiar columnas excedentes
        If Not fun825_LimpiarColumnasExcedentes(wsEnvio, vColumna_Final_HojaComprobacion) Then
            fun801_LogMessage "Advertencia: Error al limpiar columnas excedentes", False, "", strHojaEnvio
        End If
        
        fun801_LogMessage "Copia completa y limpieza completadas", False, "", strFuncion
    End If
    
    '<<<<<
    End If 'vEnabled_Parts Then
    
    '--------------------------------------------------------------------------
    ' 6. Verificar integridad de la operación
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operación...", False, "", strFuncion
    
    ' Verificación básica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificación completada: ambas hojas están vacías (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificación", False, "", strFuncion
    Else
        fun801_LogMessage "Verificación completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Comprobar cada celda y etiquetar en color Verde o ROJO cada línea
    '--------------------------------------------------------------------------
    lngLineaError = 125
    fun801_LogMessage "Comprobando valores cargados en HFM" & vbCrLf & "vs valores que pretendiamos cargar en HFM ...", False, "", strFuncion
    
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    Dim vScenario As Variant
    
    'Otras variables
    Dim vColumnaEtiqueta As Integer
    vColumnaEtiqueta = 1
    Dim vValorEnviado, vValorPretendido As Double
    Dim vValorEtiqueta As String
    Dim vEtiquetaInicial, vEtiquetaOK, vEtiquetaERROR As String
    vEtiquetaInicial = "ok": vEtiquetaOK = "ok": vEtiquetaERROR = "ERROR---ERROR"
    
    
    Application.ScreenUpdating = False
    
    
    For r = vFila_Inicial_HojaComprobacion + 2 To vFila_Final_HojaComprobacion
        'Inicializamos el valor de la columna Etiqueta
        wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaInicial
        wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = xlColorIndexNone 'Sin color de fondo
        
        For c = vColumna_Inicial_HojaComprobacion + 11 To vColumna_Final_HojaComprobacion
            
            If IsNumeric(wsEnvio.Cells(r, c).Value) Then
                vValorEnviado = CDbl(wsEnvio.Cells(r, c).Value)
            Else
                vValorEnviado = wsEnvio.Cells(r, c).Value
            End If
            If IsNumeric(wsComprobacion.Cells(r, c).Value) Then
                vValorPretendido = CDbl(wsComprobacion.Cells(r, c).Value)
            Else
                vValorPretendido = wsComprobacion.Cells(r, c).Value
            End If
            vValorEtiqueta = wsEnvio.Cells(r, vColumnaEtiqueta).Value
            
            If vValorEtiqueta = vEtiquetaERROR Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaERROR
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(255, 99, 71) 'Red Tomato
            ElseIf vValorEnviado = vValorPretendido Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaOK
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(50, 205, 50) 'LimeGreen
                
            ElseIf Abs(vValorEnviado - vValorPretendido) < 0.000001 Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaOK
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(50, 205, 50) 'LimeGreen
            Else
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaERROR
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(255, 99, 71) 'Red Tomato
            End If
            
        Next c
    Next r
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 6.3. Coger la Entity, FY, Scenario y llevarlo a variables especificias
    '--------------------------------------------------------------------------
    lngLineaError = 126
    fun801_LogMessage "Comprobando/almacenando valores de Entity, FY, y Scenario", False, "", strFuncion
    
    'Variables para tomar el dato/miembro del POV
    'Dim vYear_xPL As String
    'Dim vScenario_xPL As String
    'Dim vEntity_xPL As String
    
    'Variables para buscar las filas/columnas necesarias
    Dim vFilaReferencia As Integer
    Dim vColumnaReferencia As Integer
    'Variables para buscar columnas especificas
    Dim vColumnaEscenario As Integer
    Dim vColumnaYear As Integer
    Dim vColumnaEntity As Integer
    
    'Inicializamos las variables que me indican numero de fila/columna
    vFilaReferencia = vFila_Inicial_HojaComprobacion + 2
    vColumnaReferencia = vColumna_Inicial_HojaComprobacion
    vColumnaEscenario = vColumnaReferencia + 0
    vColumnaYear = vColumnaReferencia + 1
    vColumnaEntity = vColumnaReferencia + 3
    
    'Tomamos el valor para el Escenario, Year, Entity
    vScenario_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaEscenario).Value
    vYear_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaYear).Value
    vEntity_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaEntity).Value
        
    '--------------------------------------------------------------------------
    ' 7. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 130
    fun801_LogMessage "Copia de datos de comprobación a envío completada con éxito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuración
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F008_Comparar_Datos_HojaComprobacion_vs_HojaEnvio = False
End Function

'Sigue aqui: 20250609
Public Function F008_Actualizar_Informe_PL_AdHoc(ByVal strHojaPLAH As String) As Boolean
    
    '******************************************************************************
    ' Detecta donde estan los datos en la hoja del Informe PL AdHoc
    ' Modifica Scenario, Year, Entity
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Parámetros:
    ' - strInformePLAH: Nombre de la hoja del Informe de PL en formato AdHoc
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    'Variable para habilitar/deshabilitar partes de esta funcion
    Dim vEnabled_Parts As Boolean
    'vEnabled_Parts = True
    'If vEnabled_Parts Then
    'End If 'vEnabled_Parts Then
    
    ' Variables para mostrar información de rangos
    Dim strMensajeRangosDeTrabajo As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas de trabajo
    Dim wsHojaPLAH As Worksheet
    
    ' Variables para rangos de la hoja de comprobación
    Dim vFila_Inicial_HojaPLAH As Long
    Dim vFila_Final_HojaPLAH As Long
    Dim vColumna_Inicial_HojaPLAH As Long
    Dim vColumna_Final_HojaPLAH As Long
        
    
    ' Inicialización
    strFuncion = "F008_Actualizar_Informe_PL_AdHoc" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F008_Actualizar_Informe_PL_AdHoc"
    F008_Actualizar_Informe_PL_AdHoc = False
    
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para comprobar datos enviados...", False, "", strFuncion
    
    ' Validar hoja de envío
    If Not fun802_SheetExists(strHojaPLAH) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de envío no existe: " & strHojaEnvio
    End If
        
    ' Obtener referencias a las hojas
    Set wsHojaPLAH = ThisWorkbook.Worksheets(strHojaPLAH)
    
    ' Verificar que las referencias son válidas
    If wsHojaPLAH Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja del Informe PL AdHoc"
    End If
        
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave específicas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento específico
    ' Solo si necesitas valores diferentes a los por defecto
    Dim vEscenarioAdmitido, vUltimoMesCarga As String
    vEscenarioAdmitido = UCase(Trim(CONST_ESCENARIO_ADMITIDO))
    vUltimoMesCarga = UCase(Trim(CONST_ULTIMO_MES_DE_CARGA))
    Call fun826_ConfigurarPalabrasClave(vEscenarioAdmitido, vEscenarioAdmitido, vEscenarioAdmitido, vUltimoMesCarga)
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Detectando rangos de datos en hoja de Informe PL AdHoc...", False, "", strHojaComprobacion
    
    If Not fun822_DetectarRangoCompletoHoja(wsHojaPLAH, _
                                           vFila_Inicial_HojaPLAH, _
                                           vFila_Final_HojaPLAH, _
                                           vColumna_Inicial_HojaPLAH, _
                                           vColumna_Final_HojaPLAH) Then
        Err.Raise ERROR_BASE_IMPORT + 705, strFuncion, _
            "Error al detectar rangos en hoja de Informe PL AdHoc"
    End If
    
    fun801_LogMessage "Rangos de hoja Informe PL AdHoc - Filas: " & vFila_Inicial_HojaPLAH & " a " & vFila_Final_HojaPLAH & _
                      ", Columnas: " & vColumna_Inicial_HojaPLAH & " a " & vColumna_Final_HojaPLAH, _
                      False, "", strHojaPLAH
    
    vFila_Inicial_HojaPLAH = vFila_Inicial_HojaPLAH - 1 'Le quitamos 1, para que considere también la fila en la que están los headers de los meses M01 ... M12
    vColumna_Final_HojaPLAH = vColumna_Final_HojaPLAH + 22
    
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar información completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then

        lngLineaError = 125
        strMensajeRangosCompleto = "INFORMACIÓN COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE ENVÍO: " & strHojaEnvio & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaEnvio & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaEnvio & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaEnvio & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaEnvio & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaEnvio - vFila_Inicial_HojaEnvio + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaEnvio - vColumna_Inicial_HojaEnvio + 1) & vbCrLf & vbCrLf
        
        MsgBox strMensajeRangosCompleto, vbInformation, "Rangos Completos - " & strFuncion
        
    End If 'vEnabled_Parts Then
    
    
    '--------------------------------------------------------------------------
    ' 5. Procesar según el resultado de la comparación
    '--------------------------------------------------------------------------
    
    
    '----------------------------------------------------------------------
    ' 5.1. Rangos iguales: Copiar datos específicos (filas+2, columnas+11)
    '----------------------------------------------------------------------
    lngLineaError = 90003
    fun801_LogMessage "Ejecutando copia específica para rangos idénticos...", False, "", strFuncion
    
    ' Validar que hay suficientes filas y columnas para el offset
    'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
       (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
        
        ' Definir rango origen (desde comprobación)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino (hacia envío)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores únicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                "Error al copiar valores específicos"
        End If
        
        fun801_LogMessage "Copia específica completada correctamente", False, "", strFuncion
    'Else
    '    fun801_LogMessage "Advertencia: Offset insuficiente para copia específica, omitiendo operación", False, "", strFuncion
    'End If
    
        
    '--------------------------------------------------------------------------
    ' 6. Verificar integridad de la operación
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operación...", False, "", strFuncion
    
    ' Verificación básica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificación completada: ambas hojas están vacías (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificación", False, "", strFuncion
    Else
        fun801_LogMessage "Verificación completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Comprobar cada celda y etiquetar en color Verde o ROJO cada línea
    '--------------------------------------------------------------------------
    lngLineaError = 125
    fun801_LogMessage "Comprobando valores cargados en HFM" & vbCrLf & "vs valores que pretendiamos cargar en HFM ...", False, "", strFuncion
    
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    Dim vScenario As Variant
    
    'Otras variables
    Dim vColumnaEtiqueta As Integer
    vColumnaEtiqueta = 1
    Dim vValorEnviado, vValorPretendido As Double
    Dim vValorEtiqueta As String
    Dim vEtiquetaInicial, vEtiquetaOK, vEtiquetaERROR As String
    vEtiquetaInicial = "ok": vEtiquetaOK = "ok": vEtiquetaERROR = "ERROR---ERROR"
    
    
    Application.ScreenUpdating = False
    
    
    For r = vFila_Inicial_HojaComprobacion + 2 To vFila_Final_HojaComprobacion
        'Inicializamos el valor de la columna Etiqueta
        wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaInicial
        wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = xlColorIndexNone 'Sin color de fondo
        
        For c = vColumna_Inicial_HojaComprobacion + 11 To vColumna_Final_HojaComprobacion
            
            If IsNumeric(wsEnvio.Cells(r, c).Value) Then
                vValorEnviado = CDbl(wsEnvio.Cells(r, c).Value)
            Else
                vValorEnviado = wsEnvio.Cells(r, c).Value
            End If
            If IsNumeric(wsComprobacion.Cells(r, c).Value) Then
                vValorPretendido = CDbl(wsComprobacion.Cells(r, c).Value)
            Else
                vValorPretendido = wsComprobacion.Cells(r, c).Value
            End If
            vValorEtiqueta = wsEnvio.Cells(r, vColumnaEtiqueta).Value
            
            If vValorEtiqueta = vEtiquetaERROR Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaERROR
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(255, 99, 71) 'Red Tomato
            ElseIf vValorEnviado = vValorPretendido Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaOK
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(50, 205, 50) 'LimeGreen
                
            ElseIf Abs(vValorEnviado - vValorPretendido) < 0.000001 Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaOK
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(50, 205, 50) 'LimeGreen
            Else
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaERROR
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(255, 99, 71) 'Red Tomato
            End If
            
        Next c
    Next r
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 6.3. Coger la Entity, FY, Scenario y llevarlo a variables especificias
    '--------------------------------------------------------------------------
    lngLineaError = 126
    fun801_LogMessage "Comprobando/almacenando valores de Entity, FY, y Scenario", False, "", strFuncion
    
    'Variables para tomar el dato/miembro del POV
    'Dim vYear_xPL As String
    'Dim vScenario_xPL As String
    'Dim vEntity_xPL As String
    
    'Variables para buscar las filas/columnas necesarias
    Dim vFilaReferencia As Integer
    Dim vColumnaReferencia As Integer
    'Variables para buscar columnas especificas
    Dim vColumnaEscenario As Integer
    Dim vColumnaYear As Integer
    Dim vColumnaEntity As Integer
    
    'Inicializamos las variables que me indican numero de fila/columna
    vFilaReferencia = vFila_Inicial_HojaComprobacion + 2
    vColumnaReferencia = vColumna_Inicial_HojaComprobacion
    vColumnaEscenario = vColumnaReferencia + 0
    vColumnaYear = vColumnaReferencia + 1
    vColumnaEntity = vColumnaReferencia + 3
    
    'Tomamos el valor para el Escenario, Year, Entity
    vScenario_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaEscenario).Value
    vYear_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaYear).Value
    vEntity_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaEntity).Value
        
    '--------------------------------------------------------------------------
    ' 7. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 130
    fun801_LogMessage "Copia de datos de comprobación a envío completada con éxito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F008_Actualizar_Informe_PL_AdHoc = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuración
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F008_Actualizar_Informe_PL_AdHoc = False
End Function





