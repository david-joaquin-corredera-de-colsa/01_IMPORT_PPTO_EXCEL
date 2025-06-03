Attribute VB_Name = "Modulo_012_FUNC_Principales_02"
Option Explicit
Public Function F004_Forzar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCIÓN: F004_Forzar_Delimitadores_en_Excel
    ' PROPÓSITO: Fuerza los delimitadores decimal y de miles en Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' PARÁMETROS: Ninguno
    ' RETORNA: Boolean (True = éxito, False = error)
    '
    ' RESUMEN DE PASOS:
    ' 1. Inicialización de variables globales si están vacías
    ' 2. Verificación de compatibilidad del sistema
    ' 3. Backup de configuración actual del usuario
    ' 4. Aplicación de nuevos delimitadores usando Application.International
    ' 5. Verificación de aplicación correcta
    ' 6. Manejo exhaustivo de errores con información detallada
    ' 7. Retorno de estado de éxito/fallo
    ' =============================================================================

    ' Variables de control de errores
    Dim strFuncionActual As String
    Dim strTipoError As String
    Dim lngLineaError As Long
    
    ' Variables de trabajo
    Dim strDelimitadorDecimalAnterior As String
    Dim strDelimitadorMilesAnterior As String
    Dim blnConfiguracionCambiada As Boolean
    
    ' Inicialización
    strFuncionActual = "F004_Forzar_Delimitadores_en_Excel"
    F004_Forzar_Delimitadores_en_Excel = False
    blnConfiguracionCambiada = False
    
    On Error GoTo ErrorHandler
    
    ' =========================================================================
    ' PASO 1: Inicialización de variables globales
    ' =========================================================================
    lngLineaError = 50
    Call fun801_InicializarVariablesGlobales
    
    ' =========================================================================
    ' PASO 2: Verificación de compatibilidad
    ' =========================================================================
    lngLineaError = 60
    If Not fun802_VerificarCompatibilidad() Then
        strTipoError = "Error de compatibilidad del sistema"
        GoTo ErrorHandler
    End If
    
    ' =========================================================================
    ' PASO 3: Backup de configuración actual
    ' =========================================================================
    lngLineaError = 70
    Call fun803_ObtenerConfiguracionActual(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
    
    ' =========================================================================
    ' PASO 4: Aplicación de nuevos delimitadores
    ' =========================================================================
    lngLineaError = 80
    If fun804_AplicarNuevosDelimitadores() Then
        blnConfiguracionCambiada = True
        
        ' =====================================================================
        ' PASO 5: Verificación de aplicación correcta
        ' =====================================================================
        lngLineaError = 90
        If fun805_VerificarAplicacionDelimitadores() Then
            F004_Forzar_Delimitadores_en_Excel = True
        Else
            strTipoError = "Error en verificación de delimitadores aplicados"
            GoTo ErrorHandler
        End If
    Else
        strTipoError = "Error al aplicar nuevos delimitadores"
        GoTo ErrorHandler
    End If
    
    Exit Function

' =============================================================================
' CONTROL DE ERRORES EXHAUSTIVO
' =============================================================================
ErrorHandler:
    ' Restaurar configuración anterior si se cambió
    If blnConfiguracionCambiada Then
        On Error Resume Next
        Call fun806_RestaurarConfiguracion(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
        On Error GoTo 0
    End If
    
    ' Mostrar información detallada del error
    Call fun807_MostrarErrorDetallado(strFuncionActual, strTipoError, lngLineaError, Err.Number, Err.Description)
    
    F004_Forzar_Delimitadores_en_Excel = False
End Function


Public Function F004_Restaurar_Delimitadores_en_Excel() As Boolean

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
    ' 10. Verificar variable global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    ' 11. Si es True, ocultar la hoja de delimitadores al finalizar
    ' 12. Manejo exhaustivo de errores con información detallada y número de línea
    '
    ' Parámetros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
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
        vHojaDelimitadoresExcelOriginales = CONST_HOJA_DELIMITADORES_ORIGINALES
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
    
    ' Verificar la variable global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    If CONST_OCULTAR_REPOSITORIO_DELIMITADORES = True Then
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
    strFuncion = "F005_Procesar_Hoja_Comprobacion"
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
    lngLineaError = 90
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



Public Function F003_Procesar_Hoja_Envio(ByVal strHojaWorking As String, _
                                         ByVal strHojaEnvio As String) As Boolean
    
    '******************************************************************************
    ' FUNCI?N PRINCIPAL MEJORADA: F003_Procesar_Hoja_Envio
    ' Fecha y Hora de Creaci?n Original: 2025-06-01 19:20:05 UTC
    ' Fecha y Hora de Modificaci?n: 2025-06-02 03:27:31 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar parámetros y obtener referencias a hojas
    ' 2. Copiar contenido de hoja Working a hoja de Envío
    ' 3. Detectar rangos de datos en hoja de envío
    ' 4. Calcular variables de columnas de control
    ' 5. Mostrar información de variables (opcional)
    ' 6. Borrar contenido de columnas innecesarias
    ' 7. Filtrar líneas basado en criterios específicos
    ' 8. NUEVO: Borrar contenido y formatos de columna vColumna_LineaSuma
    ' 9. NUEVO: Detectar primera fila con contenido después de limpieza
    ' 10. NUEVO: Añadir headers de columnas identificativas (fila -1)
    ' 11. NUEVO: Añadir headers de meses (fila -2)
    ' 12. Proceso completado exitosamente
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas
    Dim wsWorking As Worksheet
    Dim wsEnvio As Worksheet
    
    ' Variables para rangos de datos
    Dim vFila_Inicial As Long
    Dim vFila_Final As Long
    Dim vColumna_Inicial As Long
    Dim vColumna_Final As Long
    
    ' Variables para columnas de control
    Dim vColumna_IdentificadorDeLinea As Long
    Dim vColumna_LineaRepetida As Long
    Dim vColumna_LineaTratada As Long
    Dim vColumna_LineaSuma As Long
    
    ' NUEVAS VARIABLES para funcionalidad adicional
    Dim vFila_Inicial_HojaLimpia As Long
    
    ' Inicialización
    strFuncion = "F003_Procesar_Hoja_Envio"
    F003_Procesar_Hoja_Envio = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros y obtener referencias a hojas
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas de trabajo...", False, "", strFuncion
    
    If Not fun802_SheetExists(strHojaWorking) Then
        Err.Raise ERROR_BASE_IMPORT + 101, strFuncion, _
            "La hoja de trabajo no existe: " & strHojaWorking
    End If
    
    If Not fun802_SheetExists(strHojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 102, strFuncion, _
            "La hoja de envío no existe: " & strHojaEnvio
    End If
    
    Set wsWorking = ThisWorkbook.Worksheets(strHojaWorking)
    Set wsEnvio = ThisWorkbook.Worksheets(strHojaEnvio)
    
    '--------------------------------------------------------------------------
    ' 2. Copiar contenido de hoja Working a hoja de Envío
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Copiando contenido de hoja Working a hoja de Envío...", False, "", strFuncion
    
    If Not fun812_CopiarContenidoCompleto(wsWorking, wsEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 103, strFuncion, _
            "Error al copiar contenido entre hojas"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Detectar rangos de datos en hoja de envío
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Detectando rangos de datos en hoja de envío...", False, "", strFuncion
    
    If Not fun813_DetectarRangoCompleto(wsEnvio, vFila_Inicial, vFila_Final, _
                                       vColumna_Inicial, vColumna_Final) Then
        Err.Raise ERROR_BASE_IMPORT + 104, strFuncion, _
            "Error al detectar rangos de datos"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Calcular variables de columnas de control
    '--------------------------------------------------------------------------
    lngLineaError = 90
    fun801_LogMessage "Calculando variables de columnas de control...", False, "", strFuncion
    
    vColumna_IdentificadorDeLinea = vColumna_Inicial + 23
    vColumna_LineaRepetida = vColumna_Inicial + 24
    vColumna_LineaTratada = vColumna_Inicial + 25
    vColumna_LineaSuma = vColumna_Inicial + 26
    
    ' Mostrar información de variables (activar/desactivar cambiando True/False)
    
    If True Then ' Cambiar a False para desactivar el mensaje
        If CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS Then Call fun814_MostrarInformacionColumnas(vColumna_Inicial, vColumna_Final, _
                                              vColumna_IdentificadorDeLinea, _
                                              vColumna_LineaRepetida, _
                                              vColumna_LineaTratada, _
                                              vColumna_LineaSuma, _
                                              vFila_Inicial, vFila_Final)
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Borrar contenido de columnas innecesarias
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Borrando contenido de columnas innecesarias...", False, "", strFuncion
    
    If Not fun815_BorrarColumnasInnecesarias(wsEnvio, vFila_Inicial, vFila_Final, _
                                            vColumna_Inicial, vColumna_IdentificadorDeLinea, _
                                            vColumna_LineaRepetida, vColumna_LineaSuma) Then
        Err.Raise ERROR_BASE_IMPORT + 105, strFuncion, _
            "Error al borrar columnas innecesarias"
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Filtrar líneas basado en criterios específicos
    '--------------------------------------------------------------------------
    lngLineaError = 110
    fun801_LogMessage "Filtrando líneas basado en criterios específicos...", False, "", strFuncion
    
    If Not fun816_FiltrarLineasEspecificas(wsEnvio, vFila_Inicial, vFila_Final, _
                                          vColumna_Inicial, vColumna_LineaTratada) Then
        Err.Raise ERROR_BASE_IMPORT + 106, strFuncion, _
            "Error al filtrar líneas específicas"
    End If
    
    '--------------------------------------------------------------------------
    ' 7. NUEVA FUNCIONALIDAD: Borrar contenido y formatos de columna vColumna_LineaSuma
    '--------------------------------------------------------------------------
    lngLineaError = 115
    fun801_LogMessage "Borrando contenido y formatos de columna LineaSuma...", False, "", strFuncion
    
    If Not fun818_BorrarColumnaLineaSuma(wsEnvio, vColumna_LineaSuma) Then
        Err.Raise ERROR_BASE_IMPORT + 107, strFuncion, _
            "Error al borrar columna LineaSuma"
    End If
    
    '--------------------------------------------------------------------------
    ' 8. NUEVA FUNCIONALIDAD: Detectar primera fila con contenido después de limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 118
    fun801_LogMessage "Detectando primera fila con contenido después de limpieza...", False, "", strFuncion
    
    If Not fun819_DetectarPrimeraFilaContenido(wsEnvio, vColumna_Inicial, vFila_Inicial_HojaLimpia) Then
        Err.Raise ERROR_BASE_IMPORT + 108, strFuncion, _
            "Error al detectar primera fila con contenido"
    End If
    
    fun801_LogMessage "Primera fila con contenido detectada: " & vFila_Inicial_HojaLimpia, False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 9. NUEVA FUNCIONALIDAD: Añadir headers de columnas identificativas
    '--------------------------------------------------------------------------
    lngLineaError = 121
    fun801_LogMessage "Añadiendo headers de columnas identificativas...", False, "", strFuncion
    
    If Not fun820_AnadirHeadersIdentificativos(wsEnvio, vFila_Inicial_HojaLimpia, vColumna_Inicial) Then
        Err.Raise ERROR_BASE_IMPORT + 109, strFuncion, _
            "Error al añadir headers identificativos"
    End If
    
    '--------------------------------------------------------------------------
    ' 10. NUEVA FUNCIONALIDAD: Añadir headers de meses
    '--------------------------------------------------------------------------
    lngLineaError = 124
    fun801_LogMessage "Añadiendo headers de meses...", False, "", strFuncion
    
    If Not fun821_AnadirHeadersMeses(wsEnvio, vFila_Inicial_HojaLimpia, vColumna_Inicial) Then
        Err.Raise ERROR_BASE_IMPORT + 110, strFuncion, _
            "Error al añadir headers de meses"
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Proceso completado exitosamente
    '--------------------------------------------------------------------------
    lngLineaError = 127
    fun801_LogMessage "Procesamiento de hoja de envío completado correctamente", False, "", strFuncion
    
    F003_Procesar_Hoja_Envio = True
    Exit Function

GestorErrores:
    ' Construción del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F003_Procesar_Hoja_Envio = False
End Function

Public Function F007_Copiar_Datos_de_Comprobacion_a_Envio(ByVal strHojaEnvio As String, _
                                                          ByVal strHojaComprobacion As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F007_Copiar_Datos_de_Comprobacion_a_Envio
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
    strFuncion = "F007_Copiar_Datos_de_Comprobacion_a_Envio"
    F007_Copiar_Datos_de_Comprobacion_a_Envio = False
    lngLineaError = 0
    vLosRangosSonIguales = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para copia de comprobación a envío...", False, "", strFuncion
    
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
    Call fun826_ConfigurarPalabrasClave("BUDGET_OS", "BUDGET_OS", "BUDGET_OS", "M12")
    
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
    
    vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaComprobacion - 1
    'vFila_Final_HojaComprobacion = 119
    'vColumna_Inicial_HojaComprobacion = 2
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
            
    vFila_Inicial_HojaEnvio = vFila_Inicial_HojaEnvio - 1
    'vFila_Final_HojaEnvio = 119
    'vColumna_Inicial_HojaEnvio = 2
    vColumna_Final_HojaEnvio = vColumna_Inicial_HojaEnvio + 22
            
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar información completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    lngLineaError = 125
    'Dim strMensajeRangosCompleto As String
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
    vLosRangosSonIguales = True
    '--------------------------------------------------------------------------
    ' 5. Procesar según el resultado de la comparación
    '--------------------------------------------------------------------------
    If vLosRangosSonIguales = True Then
        '----------------------------------------------------------------------
        ' 5.1. Rangos iguales: Copiar datos específicos (filas+2, columnas+11)
        '----------------------------------------------------------------------
        lngLineaError = 90
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
        ' 5.2. Rangos diferentes: Copiar contenido completo y limpiar excesos
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
    ' 6.1. Verificar integridad de la operación
    '--------------------------------------------------------------------------
    lngLineaError = 125
    fun801_LogMessage "Editando cada celda del rango para poder hacer Submit...", False, "", strFuncion
    
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    Dim vScenario As Variant
    
    Application.ScreenUpdating = False
    
    For r = vFila_Inicial_HojaComprobacion + 2 To vFila_Final_HojaComprobacion
        For c = vColumna_Inicial_HojaComprobacion + 11 To vColumna_Final_HojaComprobacion
            vScenario = Trim(Cells(r, vColumna_Inicial_HojaComprobacion).Value)
            If vScenario <> "" Then
                vValor = Cells(r, c).Value
                Cells(r, c).Value = vValor
            End If
        Next c
    Next r
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 7. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 130
    fun801_LogMessage "Copia de datos de comprobación a envío completada con éxito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F007_Copiar_Datos_de_Comprobacion_a_Envio = True
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
    F007_Copiar_Datos_de_Comprobacion_a_Envio = False
End Function

Public Function F008_Ocultar_Hojas_Antiguas() As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F008_Ocultar_Hojas_Antiguas
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Fecha y Hora de Modificación: 2025-06-03 04:36:36 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Función para gestionar la visibilidad de hojas en el libro, ocultando hojas
    ' antiguas de importación y manteniendo visibles solo las más recientes.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que existan hojas en el libro
    ' 2. Procesar hojas del sistema (sin modificar su visibilidad)
    ' 3. Ocultar hojas de importación, working y comprobación
    ' 4. Identificar hojas de envío por fecha y hora
    ' 5. Si hay menos de 4 hojas Import_Envio_: mantener todas visibles
    ' 6. Si hay 4 o más hojas Import_Envio_: mantener visibles solo las 3 más recientes
    ' 7. Registrar resultados en el log del sistema
    '
    ' Parámetros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento de hojas
    Dim ws As Worksheet
    Dim i As Long
    Dim intTotalHojas As Integer
    
    ' Variables para contadores
    Dim intHojasSistema As Integer
    Dim intHojasImportOcultadas As Integer
    Dim intHojasEnvioEncontradas As Integer
    Dim intHojasEnvioOcultadas As Integer
    
    ' Inicialización
    strFuncion = "F008_Ocultar_Hojas_Antiguas"
    F008_Ocultar_Hojas_Antiguas = False
    lngLineaError = 0
    intHojasSistema = 0
    intHojasImportOcultadas = 0
    intHojasEnvioEncontradas = 0
    intHojasEnvioOcultadas = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que existan hojas en el libro
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando proceso de ocultación de hojas antiguas...", False, "", strFuncion
    
    intTotalHojas = ThisWorkbook.Worksheets.Count
    If intTotalHojas = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "El libro no contiene hojas de trabajo"
    End If
    
    fun801_LogMessage "Total de hojas en el libro: " & intTotalHojas, False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 2. Procesar hojas del sistema (sin modificar visibilidad)
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Procesando hojas del sistema...", False, "", strFuncion
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        
        If fun821_EsHojaSistema(ws.Name) Then
            intHojasSistema = intHojasSistema + 1
            fun801_LogMessage "Hoja del sistema encontrada (sin modificar): " & ws.Name, _
                              False, "", strFuncion
        End If
    Next i
    
    '--------------------------------------------------------------------------
    ' 3. Ocultar hojas de importación, working y comprobación (NO las de envío)
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Ocultando hojas de importación, working y comprobación...", _
                      False, "", strFuncion
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        
        If fun822_EsHojaImportacionSinEnvio(ws.Name) Then
            If fun823_OcultarHojaSiVisible(ws) Then
                intHojasImportOcultadas = intHojasImportOcultadas + 1
                fun801_LogMessage "Hoja ocultada: " & ws.Name, False, "", strFuncion
            End If
        End If
    Next i
    
    '--------------------------------------------------------------------------
    ' 4. Contar hojas de envío
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Contando hojas de envío...", False, "", strFuncion
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        
        If fun824_EsHojaEnvio(ws.Name) Then
            intHojasEnvioEncontradas = intHojasEnvioEncontradas + 1
        End If
    Next i
    
    fun801_LogMessage "Hojas de envío encontradas: " & intHojasEnvioEncontradas, _
                      False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 5. Procesar hojas de envío según la cantidad encontrada
    '--------------------------------------------------------------------------
    lngLineaError = 90
    If intHojasEnvioEncontradas < 4 Then
        fun801_LogMessage "Menos de 4 hojas de envío: manteniendo todas visibles", _
                          False, "", strFuncion
        intHojasEnvioOcultadas = 0
    Else
        fun801_LogMessage "4 o más hojas de envío: manteniendo visibles las 3 más recientes", _
                          False, "", strFuncion
        
        If Not fun825_ProcesarHojasEnvioCorregido(intHojasEnvioOcultadas) Then
            Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
                "Error al procesar hojas de envío"
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Registrar resultados finales
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Proceso completado - Resumen:", False, "", strFuncion
    fun801_LogMessage "- Hojas del sistema procesadas: " & intHojasSistema, _
                      False, "", strFuncion
    fun801_LogMessage "- Hojas de importación ocultadas: " & intHojasImportOcultadas, _
                      False, "", strFuncion
    fun801_LogMessage "- Hojas de envío encontradas: " & intHojasEnvioEncontradas, _
                      False, "", strFuncion
    fun801_LogMessage "- Hojas de envío ocultadas: " & intHojasEnvioOcultadas, _
                      False, "", strFuncion
    
    F008_Ocultar_Hojas_Antiguas = True
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F008_Ocultar_Hojas_Antiguas = False
End Function
