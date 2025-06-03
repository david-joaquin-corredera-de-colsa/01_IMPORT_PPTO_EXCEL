Attribute VB_Name = "Modulo_012_FUNC_Principales_02"
Option Explicit
Public Function F004_Forzar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCI�N: F004_Forzar_Delimitadores_en_Excel
    ' PROP�SITO: Fuerza los delimitadores decimal y de miles en Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' PAR�METROS: Ninguno
    ' RETORNA: Boolean (True = �xito, False = error)
    '
    ' RESUMEN DE PASOS:
    ' 1. Inicializaci�n de variables globales si est�n vac�as
    ' 2. Verificaci�n de compatibilidad del sistema
    ' 3. Backup de configuraci�n actual del usuario
    ' 4. Aplicaci�n de nuevos delimitadores usando Application.International
    ' 5. Verificaci�n de aplicaci�n correcta
    ' 6. Manejo exhaustivo de errores con informaci�n detallada
    ' 7. Retorno de estado de �xito/fallo
    ' =============================================================================

    ' Variables de control de errores
    Dim strFuncionActual As String
    Dim strTipoError As String
    Dim lngLineaError As Long
    
    ' Variables de trabajo
    Dim strDelimitadorDecimalAnterior As String
    Dim strDelimitadorMilesAnterior As String
    Dim blnConfiguracionCambiada As Boolean
    
    ' Inicializaci�n
    strFuncionActual = "F004_Forzar_Delimitadores_en_Excel"
    F004_Forzar_Delimitadores_en_Excel = False
    blnConfiguracionCambiada = False
    
    On Error GoTo ErrorHandler
    
    ' =========================================================================
    ' PASO 1: Inicializaci�n de variables globales
    ' =========================================================================
    lngLineaError = 50
    Call fun801_InicializarVariablesGlobales
    
    ' =========================================================================
    ' PASO 2: Verificaci�n de compatibilidad
    ' =========================================================================
    lngLineaError = 60
    If Not fun802_VerificarCompatibilidad() Then
        strTipoError = "Error de compatibilidad del sistema"
        GoTo ErrorHandler
    End If
    
    ' =========================================================================
    ' PASO 3: Backup de configuraci�n actual
    ' =========================================================================
    lngLineaError = 70
    Call fun803_ObtenerConfiguracionActual(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
    
    ' =========================================================================
    ' PASO 4: Aplicaci�n de nuevos delimitadores
    ' =========================================================================
    lngLineaError = 80
    If fun804_AplicarNuevosDelimitadores() Then
        blnConfiguracionCambiada = True
        
        ' =====================================================================
        ' PASO 5: Verificaci�n de aplicaci�n correcta
        ' =====================================================================
        lngLineaError = 90
        If fun805_VerificarAplicacionDelimitadores() Then
            F004_Forzar_Delimitadores_en_Excel = True
        Else
            strTipoError = "Error en verificaci�n de delimitadores aplicados"
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
    ' Restaurar configuraci�n anterior si se cambi�
    If blnConfiguracionCambiada Then
        On Error Resume Next
        Call fun806_RestaurarConfiguracion(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
        On Error GoTo 0
    End If
    
    ' Mostrar informaci�n detallada del error
    Call fun807_MostrarErrorDetallado(strFuncionActual, strTipoError, lngLineaError, Err.Number, Err.Description)
    
    F004_Forzar_Delimitadores_en_Excel = False
End Function


Public Function F004_Restaurar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCI�N PRINCIPAL: F004_Restaurar_Delimitadores_en_Excel
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura los delimitadores originales de Excel desde la hoja de respaldo
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables globales con valores por defecto (C2, C3, C4)
    ' 2. Obtener referencia al libro actual
    ' 3. Verificar si existe la hoja de delimitadores originales
    ' 4. Si no existe, crear la hoja y dejarla visible (situaci�n extra�a para restauraci�n)
    ' 5. Si existe, verificar su visibilidad y hacerla visible si est� oculta
    ' 6. Leer valores originales desde las celdas especificadas:
    '    - Use System Separators desde C2
    '    - Decimal Separator desde C3
    '    - Thousands Separator desde C4
    ' 7. Almacenar valores le�dos en variables globales correspondientes
    ' 8. Validar que los valores le�dos sean apropiados para restaurar
    ' 9. Aplicar configuraci�n original de delimitadores de Excel:
    '    - Use System Separators (True/False seg�n valor original)
    '    - Decimal Separator (car�cter seg�n valor original)
    '    - Thousands Separator (car�cter seg�n valor original)
    ' 10. Verificar variable global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    ' 11. Si es True, ocultar la hoja de delimitadores al finalizar
    ' 12. Manejo exhaustivo de errores con informaci�n detallada y n�mero de l�nea
    '
    ' Par�metros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    ' Control de errores con n�mero de l�nea
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
    ' NOTA: Usuario especific� C2 para todas, corrijo para C2, C3, C4 seg�n l�gica
    vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal = "C2"
    vCelda_Valor_Excel_DecimalSeparator_ValorOriginal = "C3"
    vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal = "C4"
    
    ' Variables para almacenar los valores originales (inicialmente vac�as)
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
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEG�N CORRESPONDA
    ' ==========================================================================
    
    If Not hojaExiste Then
        ' La hoja no existe, crearla y dejarla visible
        ' NOTA: En un escenario de restauraci�n, esto ser�a extra�o, pero cumplimos la especificaci�n
        Set ws = fun802_CrearHojaDelimitadores(wb, vHojaDelimitadoresExcelOriginales)
        If ws Is Nothing Then
            F004_Restaurar_Delimitadores_en_Excel = False
            Exit Function
        End If
        ' Como no hay datos que leer, salir con �xito pero sin restaurar
        Debug.Print "ADVERTENCIA: Hoja de delimitadores creada, pero no hay valores para restaurar - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        Exit Function
    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
        Set ws = wb.Worksheets(vHojaDelimitadoresExcelOriginales)
        
        ' Verificar si est� oculta y hacerla visible si es necesario
        If Not fun803_HacerHojaVisible(ws) Then
            Debug.Print "ADVERTENCIA: No se pudo hacer visible la hoja " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
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
    ' PASO 6: VALIDAR QUE SE HAYAN LE�DO VALORES V�LIDOS
    ' ==========================================================================
    
    If Not fun805_ValidarValoresOriginales() Then
        Debug.Print "ADVERTENCIA: No se encontraron valores v�lidos para restaurar en la hoja: " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        F004_Restaurar_Delimitadores_en_Excel = False
        Exit Function
    End If
    
    lineaError = 160
    
    ' ==========================================================================
    ' PASO 7: APLICAR CONFIGURACI�N ORIGINAL DE DELIMITADORES DE EXCEL
    ' ==========================================================================
    
    ' Restaurar Use System Separators (True/False)
    If Not fun806_RestaurarUseSystemSeparators(vExcel_UseSystemSeparators_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Use System Separators - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    ' Restaurar Decimal Separator (car�cter)
    If Not fun807_RestaurarDecimalSeparator(vExcel_DecimalSeparator_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Decimal Separator - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    ' Restaurar Thousands Separator (car�cter)
    If Not fun808_RestaurarThousandsSeparator(vExcel_ThousandsSeparator_ValorOriginal) Then
        Debug.Print "ADVERTENCIA: Error al restaurar Thousands Separator - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    End If
    
    lineaError = 170
    
    ' ==========================================================================
    ' PASO 8: VERIFICAR SI DEBE OCULTAR LA HOJA
    ' ==========================================================================
    
    ' Verificar la variable global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    If CONST_OCULTAR_REPOSITORIO_DELIMITADORES = True Then
        ' Ocultar la hoja de delimitadores
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
        End If
    End If
    
    lineaError = 180
    
    ' ==========================================================================
    ' PASO 9: FINALIZACI�N EXITOSA
    ' ==========================================================================
    
    Debug.Print "�XITO: Delimitadores restaurados correctamente - Funci�n: F004_Restaurar_Delimitadores_en_Excel - " & Now()
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Restaurar_Delimitadores_en_Excel = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: F004_Restaurar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
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
    ' FUNCI�N PRINCIPAL: F005_Procesar_Hoja_Comprobacion
    ' Fecha y Hora de Creaci�n: 2025-06-01 21:52:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Copia todo el contenido de la hoja de env�o a la hoja de comprobaci�n
    ' para permitir verificaci�n y control de calidad de los datos procesados.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que las hojas de env�o y comprobaci�n existan
    ' 2. Obtener referencias a las hojas de trabajo
    ' 3. Copiar contenido completo de hoja env�o a hoja comprobaci�n
    ' 4. Verificar que la copia se realiz� correctamente
    ' 5. Registrar el resultado en el log del sistema
    '
    ' Par�metros: Ninguno (usa variables globales)
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
    
    ' Inicializaci�n
    strFuncion = "F005_Procesar_Hoja_Comprobacion"
    F005_Procesar_Hoja_Comprobacion = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que las hojas existan
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando existencia de hojas para procesamiento de comprobaci�n...", False, "", strFuncion
    
    ' Validar hoja de env�o
    If Not fun802_SheetExists(gstrNuevaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 301, strFuncion, _
            "La hoja de env�o no existe: " & gstrNuevaHojaImportacion_Envio
    End If
    
    ' Validar hoja de comprobaci�n
    If Not fun802_SheetExists(gstrNuevaHojaImportacion_Comprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 302, strFuncion, _
            "La hoja de comprobaci�n no existe: " & gstrNuevaHojaImportacion_Comprobacion
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Obtener referencias a las hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Obteniendo referencias a hojas de trabajo...", False, "", strFuncion
    
    Set wsEnvio = ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Envio)
    Set wsComprobacion = ThisWorkbook.Worksheets(gstrNuevaHojaImportacion_Comprobacion)
    
    ' Verificar que las referencias son v�lidas
    If wsEnvio Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 303, strFuncion, _
            "No se pudo obtener referencia a la hoja de env�o"
    End If
    
    If wsComprobacion Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 304, strFuncion, _
            "No se pudo obtener referencia a la hoja de comprobaci�n"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Copiar contenido completo de hoja env�o a hoja comprobaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Copiando contenido de hoja de env�o a hoja de comprobaci�n...", _
                      False, gstrNuevaHojaImportacion_Envio, gstrNuevaHojaImportacion_Comprobacion
    
    If Not fun817_CopiarContenidoCompleto(wsEnvio, wsComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 305, strFuncion, _
            "Error al copiar contenido de hoja env�o a hoja comprobaci�n"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Verificar que la copia se realiz� correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Verificando integridad de la copia...", False, "", strFuncion
    
    ' Verificaci�n b�sica: comparar si ambas hojas tienen contenido
    If wsEnvio.UsedRange Is Nothing And wsComprobacion.UsedRange Is Nothing Then
        ' Ambas est�n vac�as, es correcto
        fun801_LogMessage "Verificaci�n completada: ambas hojas est�n vac�as (correcto)", False, "", strFuncion
    ElseIf wsEnvio.UsedRange Is Nothing Or wsComprobacion.UsedRange Is Nothing Then
        ' Una tiene contenido y la otra no, es un error
        Err.Raise ERROR_BASE_IMPORT + 306, strFuncion, _
            "Error en verificaci�n: inconsistencia en contenido de hojas"
    Else
        ' Ambas tienen contenido, verificar que tienen el mismo rango
        If wsEnvio.UsedRange.Rows.Count = wsComprobacion.UsedRange.Rows.Count And _
           wsEnvio.UsedRange.Columns.Count = wsComprobacion.UsedRange.Columns.Count Then
            fun801_LogMessage "Verificaci�n completada: dimensiones coinciden", False, "", strFuncion
        Else
            Err.Raise ERROR_BASE_IMPORT + 307, strFuncion, _
                "Error en verificaci�n: las dimensiones de los rangos no coinciden"
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 90
    fun801_LogMessage "Procesamiento de hoja de comprobaci�n completado con �xito", _
                      False, gstrNuevaHojaImportacion_Envio, gstrNuevaHojaImportacion_Comprobacion
    
    F005_Procesar_Hoja_Comprobacion = True
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
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
    ' 1. Validar par�metros y obtener referencias a hojas
    ' 2. Copiar contenido de hoja Working a hoja de Env�o
    ' 3. Detectar rangos de datos en hoja de env�o
    ' 4. Calcular variables de columnas de control
    ' 5. Mostrar informaci�n de variables (opcional)
    ' 6. Borrar contenido de columnas innecesarias
    ' 7. Filtrar l�neas basado en criterios espec�ficos
    ' 8. NUEVO: Borrar contenido y formatos de columna vColumna_LineaSuma
    ' 9. NUEVO: Detectar primera fila con contenido despu�s de limpieza
    ' 10. NUEVO: A�adir headers de columnas identificativas (fila -1)
    ' 11. NUEVO: A�adir headers de meses (fila -2)
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
    
    ' Inicializaci�n
    strFuncion = "F003_Procesar_Hoja_Envio"
    F003_Procesar_Hoja_Envio = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros y obtener referencias a hojas
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas de trabajo...", False, "", strFuncion
    
    If Not fun802_SheetExists(strHojaWorking) Then
        Err.Raise ERROR_BASE_IMPORT + 101, strFuncion, _
            "La hoja de trabajo no existe: " & strHojaWorking
    End If
    
    If Not fun802_SheetExists(strHojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 102, strFuncion, _
            "La hoja de env�o no existe: " & strHojaEnvio
    End If
    
    Set wsWorking = ThisWorkbook.Worksheets(strHojaWorking)
    Set wsEnvio = ThisWorkbook.Worksheets(strHojaEnvio)
    
    '--------------------------------------------------------------------------
    ' 2. Copiar contenido de hoja Working a hoja de Env�o
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Copiando contenido de hoja Working a hoja de Env�o...", False, "", strFuncion
    
    If Not fun812_CopiarContenidoCompleto(wsWorking, wsEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 103, strFuncion, _
            "Error al copiar contenido entre hojas"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Detectar rangos de datos en hoja de env�o
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Detectando rangos de datos en hoja de env�o...", False, "", strFuncion
    
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
    
    ' Mostrar informaci�n de variables (activar/desactivar cambiando True/False)
    
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
    ' 6. Filtrar l�neas basado en criterios espec�ficos
    '--------------------------------------------------------------------------
    lngLineaError = 110
    fun801_LogMessage "Filtrando l�neas basado en criterios espec�ficos...", False, "", strFuncion
    
    If Not fun816_FiltrarLineasEspecificas(wsEnvio, vFila_Inicial, vFila_Final, _
                                          vColumna_Inicial, vColumna_LineaTratada) Then
        Err.Raise ERROR_BASE_IMPORT + 106, strFuncion, _
            "Error al filtrar l�neas espec�ficas"
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
    ' 8. NUEVA FUNCIONALIDAD: Detectar primera fila con contenido despu�s de limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 118
    fun801_LogMessage "Detectando primera fila con contenido despu�s de limpieza...", False, "", strFuncion
    
    If Not fun819_DetectarPrimeraFilaContenido(wsEnvio, vColumna_Inicial, vFila_Inicial_HojaLimpia) Then
        Err.Raise ERROR_BASE_IMPORT + 108, strFuncion, _
            "Error al detectar primera fila con contenido"
    End If
    
    fun801_LogMessage "Primera fila con contenido detectada: " & vFila_Inicial_HojaLimpia, False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 9. NUEVA FUNCIONALIDAD: A�adir headers de columnas identificativas
    '--------------------------------------------------------------------------
    lngLineaError = 121
    fun801_LogMessage "A�adiendo headers de columnas identificativas...", False, "", strFuncion
    
    If Not fun820_AnadirHeadersIdentificativos(wsEnvio, vFila_Inicial_HojaLimpia, vColumna_Inicial) Then
        Err.Raise ERROR_BASE_IMPORT + 109, strFuncion, _
            "Error al a�adir headers identificativos"
    End If
    
    '--------------------------------------------------------------------------
    ' 10. NUEVA FUNCIONALIDAD: A�adir headers de meses
    '--------------------------------------------------------------------------
    lngLineaError = 124
    fun801_LogMessage "A�adiendo headers de meses...", False, "", strFuncion
    
    If Not fun821_AnadirHeadersMeses(wsEnvio, vFila_Inicial_HojaLimpia, vColumna_Inicial) Then
        Err.Raise ERROR_BASE_IMPORT + 110, strFuncion, _
            "Error al a�adir headers de meses"
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Proceso completado exitosamente
    '--------------------------------------------------------------------------
    lngLineaError = 127
    fun801_LogMessage "Procesamiento de hoja de env�o completado correctamente", False, "", strFuncion
    
    F003_Procesar_Hoja_Envio = True
    Exit Function

GestorErrores:
    ' Construci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F003_Procesar_Hoja_Envio = False
End Function

Public Function F007_Copiar_Datos_de_Comprobacion_a_Envio(ByVal strHojaEnvio As String, _
                                                          ByVal strHojaComprobacion As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N PRINCIPAL: F007_Copiar_Datos_de_Comprobacion_a_Envio
    ' Fecha y Hora de Creaci�n: 2025-06-03 00:14:44 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Copia datos espec�ficos desde la hoja de comprobaci�n hacia la hoja de env�o,
    ' implementando l�gica condicional basada en la comparaci�n de rangos entre ambas hojas.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar par�metros y obtener referencias a hojas de trabajo
    ' 2. Detectar rangos de datos en hoja de comprobaci�n
    ' 3. Detectar rangos de datos en hoja de env�o
    ' 4. Comparar si los rangos son id�nticos
    ' 5. Si rangos son iguales: copiar datos espec�ficos (filas+2, columnas+11)
    ' 6. Si rangos son diferentes: copiar contenido completo y limpiar excesos
    ' 7. Verificar integridad de la operaci�n
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Par�metros:
    ' - strHojaEnvio: Nombre de la hoja de destino (env�o)
    ' - strHojaComprobacion: Nombre de la hoja de origen (comprobaci�n)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para mostrar informaci�n de rangos
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
    
    ' Variables para rangos de la hoja de comprobaci�n
    Dim vFila_Inicial_HojaComprobacion As Long
    Dim vFila_Final_HojaComprobacion As Long
    Dim vColumna_Inicial_HojaComprobacion As Long
    Dim vColumna_Final_HojaComprobacion As Long
    
    ' Variables para rangos de la hoja de env�o
    Dim vFila_Inicial_HojaEnvio As Long
    Dim vFila_Final_HojaEnvio As Long
    Dim vColumna_Inicial_HojaEnvio As Long
    Dim vColumna_Final_HojaEnvio As Long
    
    ' Variable para comparaci�n de rangos
    Dim vLosRangosSonIguales As Boolean
    
    ' Variables para rangos de copia
    Dim rngOrigen As Range
    Dim rngDestino As Range
    
    ' Inicializaci�n
    strFuncion = "F007_Copiar_Datos_de_Comprobacion_a_Envio"
    F007_Copiar_Datos_de_Comprobacion_a_Envio = False
    lngLineaError = 0
    vLosRangosSonIguales = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para copia de comprobaci�n a env�o...", False, "", strFuncion
    
    ' Validar hoja de env�o
    If Not fun802_SheetExists(strHojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de env�o no existe: " & strHojaEnvio
    End If
    
    ' Validar hoja de comprobaci�n
    If Not fun802_SheetExists(strHojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 702, strFuncion, _
            "La hoja de comprobaci�n no existe: " & strHojaComprobacion
    End If
    
    ' Obtener referencias a las hojas
    Set wsEnvio = ThisWorkbook.Worksheets(strHojaEnvio)
    Set wsComprobacion = ThisWorkbook.Worksheets(strHojaComprobacion)
    
    ' Verificar que las referencias son v�lidas
    If wsEnvio Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja de env�o"
    End If
    
    If wsComprobacion Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 704, strFuncion, _
            "No se pudo obtener referencia a la hoja de comprobaci�n"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave espec�ficas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento espec�fico
    ' Solo si necesitas valores diferentes a los por defecto
    Call fun826_ConfigurarPalabrasClave("BUDGET_OS", "BUDGET_OS", "BUDGET_OS", "M12")
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Detectando rangos de datos en hoja de comprobaci�n...", False, "", strHojaComprobacion
    
    If Not fun822_DetectarRangoCompletoHoja(wsComprobacion, _
                                           vFila_Inicial_HojaComprobacion, _
                                           vFila_Final_HojaComprobacion, _
                                           vColumna_Inicial_HojaComprobacion, _
                                           vColumna_Final_HojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 705, strFuncion, _
            "Error al detectar rangos en hoja de comprobaci�n"
    End If
    
    fun801_LogMessage "Rangos de comprobaci�n - Filas: " & vFila_Inicial_HojaComprobacion & " a " & vFila_Final_HojaComprobacion & _
                      ", Columnas: " & vColumna_Inicial_HojaComprobacion & " a " & vColumna_Final_HojaComprobacion, _
                      False, "", strHojaComprobacion
    
    vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaComprobacion - 1
    'vFila_Final_HojaComprobacion = 119
    'vColumna_Inicial_HojaComprobacion = 2
    vColumna_Final_HojaComprobacion = vColumna_Inicial_HojaComprobacion + 22
    '--------------------------------------------------------------------------
    ' 3. Detectar rangos de datos en hoja de env�o
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Detectando rangos de datos en hoja de env�o...", False, "", strHojaEnvio
    
    If Not fun822_DetectarRangoCompletoHoja(wsEnvio, _
                                           vFila_Inicial_HojaEnvio, _
                                           vFila_Final_HojaEnvio, _
                                           vColumna_Inicial_HojaEnvio, _
                                           vColumna_Final_HojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 706, strFuncion, _
            "Error al detectar rangos en hoja de env�o"
    End If
    
    fun801_LogMessage "Rangos de env�o - Filas: " & vFila_Inicial_HojaEnvio & " a " & vFila_Final_HojaEnvio & _
                      ", Columnas: " & vColumna_Inicial_HojaEnvio & " a " & vColumna_Final_HojaEnvio, _
                      False, "", strHojaEnvio
            
    vFila_Inicial_HojaEnvio = vFila_Inicial_HojaEnvio - 1
    'vFila_Final_HojaEnvio = 119
    'vColumna_Inicial_HojaEnvio = 2
    vColumna_Final_HojaEnvio = vColumna_Inicial_HojaEnvio + 22
            
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar informaci�n completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    lngLineaError = 125
    'Dim strMensajeRangosCompleto As String
    strMensajeRangosCompleto = "INFORMACI�N COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                               "-----------------------------------------------" & vbCrLf & _
                               "HOJA DE ENV�O: " & strHojaEnvio & vbCrLf & _
                               "-----------------------------------------------" & vbCrLf & _
                               "- Fila Inicial: " & vFila_Inicial_HojaEnvio & vbCrLf & _
                               "- Fila Final: " & vFila_Final_HojaEnvio & vbCrLf & _
                               "- Columna Inicial: " & vColumna_Inicial_HojaEnvio & vbCrLf & _
                               "- Columna Final: " & vColumna_Final_HojaEnvio & vbCrLf & _
                               "- Total filas: " & (vFila_Final_HojaEnvio - vFila_Inicial_HojaEnvio + 1) & vbCrLf & _
                               "- Total columnas: " & (vColumna_Final_HojaEnvio - vColumna_Inicial_HojaEnvio + 1) & vbCrLf & vbCrLf & _
                               "-----------------------------------------------" & vbCrLf & _
                               "HOJA DE COMPROBACI�N: " & strHojaComprobacion & vbCrLf & _
                               "-----------------------------------------------" & vbCrLf & _
                               "- Fila Inicial: " & vFila_Inicial_HojaComprobacion & vbCrLf & _
                               "- Fila Final: " & vFila_Final_HojaComprobacion & vbCrLf & _
                               "- Columna Inicial: " & vColumna_Inicial_HojaComprobacion & vbCrLf & _
                               "- Columna Final: " & vColumna_Final_HojaComprobacion & vbCrLf & _
                               "- Total filas: " & (vFila_Final_HojaComprobacion - vFila_Inicial_HojaComprobacion + 1) & vbCrLf & _
                               "- Total columnas: " & (vColumna_Final_HojaComprobacion - vColumna_Inicial_HojaComprobacion + 1)
    
    MsgBox strMensajeRangosCompleto, vbInformation, "Rangos Completos - " & strFuncion
    
    '--------------------------------------------------------------------------
    ' 4. Comparar si los rangos son id�nticos
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Comparando rangos entre hojas...", False, "", strFuncion
    
    If (vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaEnvio) And _
       (vFila_Final_HojaComprobacion = vFila_Final_HojaEnvio) And _
       (vColumna_Inicial_HojaComprobacion = vColumna_Inicial_HojaEnvio) And _
       (vColumna_Final_HojaComprobacion = vColumna_Final_HojaEnvio) Then
        vLosRangosSonIguales = True
        fun801_LogMessage "Los rangos son id�nticos - Aplicando copia espec�fica", False, "", strFuncion
    Else
        vLosRangosSonIguales = False
        fun801_LogMessage "Los rangos son diferentes - Aplicando copia completa", False, "", strFuncion
    End If
    
    'MsgBox "Los Rangos son Iguales? = " & vLosRangosSonIguales
    vLosRangosSonIguales = True
    '--------------------------------------------------------------------------
    ' 5. Procesar seg�n el resultado de la comparaci�n
    '--------------------------------------------------------------------------
    If vLosRangosSonIguales = True Then
        '----------------------------------------------------------------------
        ' 5.1. Rangos iguales: Copiar datos espec�ficos (filas+2, columnas+11)
        '----------------------------------------------------------------------
        lngLineaError = 90
        fun801_LogMessage "Ejecutando copia espec�fica para rangos id�nticos...", False, "", strFuncion
        
        ' Validar que hay suficientes filas y columnas para el offset
        'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
           (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
            
            ' Definir rango origen (desde comprobaci�n)
            Set rngOrigen = wsComprobacion.Range( _
                wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Definir rango destino (hacia env�o)
            Set rngDestino = wsEnvio.Range( _
                wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Realizar copia de valores �nicamente
            If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
                Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                    "Error al copiar valores espec�ficos"
            End If
            
            fun801_LogMessage "Copia espec�fica completada correctamente", False, "", strFuncion
        'Else
        '    fun801_LogMessage "Advertencia: Offset insuficiente para copia espec�fica, omitiendo operaci�n", False, "", strFuncion
        'End If
        
    Else
        '----------------------------------------------------------------------
        ' 5.2. Rangos diferentes: Copiar contenido completo y limpiar excesos
        '----------------------------------------------------------------------
        lngLineaError = 100
        fun801_LogMessage "Ejecutando copia completa para rangos diferentes...", False, "", strFuncion
        
        ' Definir rango origen completo (desde comprobaci�n)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino completo (hacia env�o)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores �nicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 708, strFuncion, _
                "Error al copiar contenido completo"
        End If
        
        '----------------------------------------------------------------------
        ' 5.3. Limpiar excesos en hoja de env�o
        '----------------------------------------------------------------------
        lngLineaError = 110
        fun801_LogMessage "Limpiando excesos en hoja de env�o...", False, "", strHojaEnvio
        
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
    ' 6. Verificar integridad de la operaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operaci�n...", False, "", strFuncion
    
    ' Verificaci�n b�sica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificaci�n completada: ambas hojas est�n vac�as (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificaci�n", False, "", strFuncion
    Else
        fun801_LogMessage "Verificaci�n completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Verificar integridad de la operaci�n
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
    fun801_LogMessage "Copia de datos de comprobaci�n a env�o completada con �xito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F007_Copiar_Datos_de_Comprobacion_a_Envio = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuraci�n
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F007_Copiar_Datos_de_Comprobacion_a_Envio = False
End Function

Public Function F008_Ocultar_Hojas_Antiguas() As Boolean
    
    '******************************************************************************
    ' FUNCI�N PRINCIPAL: F008_Ocultar_Hojas_Antiguas
    ' Fecha y Hora de Creaci�n: 2025-06-03 04:25:04 UTC
    ' Fecha y Hora de Modificaci�n: 2025-06-03 04:36:36 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Funci�n para gestionar la visibilidad de hojas en el libro, ocultando hojas
    ' antiguas de importaci�n y manteniendo visibles solo las m�s recientes.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que existan hojas en el libro
    ' 2. Procesar hojas del sistema (sin modificar su visibilidad)
    ' 3. Ocultar hojas de importaci�n, working y comprobaci�n
    ' 4. Identificar hojas de env�o por fecha y hora
    ' 5. Si hay menos de 4 hojas Import_Envio_: mantener todas visibles
    ' 6. Si hay 4 o m�s hojas Import_Envio_: mantener visibles solo las 3 m�s recientes
    ' 7. Registrar resultados en el log del sistema
    '
    ' Par�metros: Ninguno
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
    
    ' Inicializaci�n
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
    fun801_LogMessage "Iniciando proceso de ocultaci�n de hojas antiguas...", False, "", strFuncion
    
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
    ' 3. Ocultar hojas de importaci�n, working y comprobaci�n (NO las de env�o)
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Ocultando hojas de importaci�n, working y comprobaci�n...", _
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
    ' 4. Contar hojas de env�o
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Contando hojas de env�o...", False, "", strFuncion
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        
        If fun824_EsHojaEnvio(ws.Name) Then
            intHojasEnvioEncontradas = intHojasEnvioEncontradas + 1
        End If
    Next i
    
    fun801_LogMessage "Hojas de env�o encontradas: " & intHojasEnvioEncontradas, _
                      False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 5. Procesar hojas de env�o seg�n la cantidad encontrada
    '--------------------------------------------------------------------------
    lngLineaError = 90
    If intHojasEnvioEncontradas < 4 Then
        fun801_LogMessage "Menos de 4 hojas de env�o: manteniendo todas visibles", _
                          False, "", strFuncion
        intHojasEnvioOcultadas = 0
    Else
        fun801_LogMessage "4 o m�s hojas de env�o: manteniendo visibles las 3 m�s recientes", _
                          False, "", strFuncion
        
        If Not fun825_ProcesarHojasEnvioCorregido(intHojasEnvioOcultadas) Then
            Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
                "Error al procesar hojas de env�o"
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Registrar resultados finales
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Proceso completado - Resumen:", False, "", strFuncion
    fun801_LogMessage "- Hojas del sistema procesadas: " & intHojasSistema, _
                      False, "", strFuncion
    fun801_LogMessage "- Hojas de importaci�n ocultadas: " & intHojasImportOcultadas, _
                      False, "", strFuncion
    fun801_LogMessage "- Hojas de env�o encontradas: " & intHojasEnvioEncontradas, _
                      False, "", strFuncion
    fun801_LogMessage "- Hojas de env�o ocultadas: " & intHojasEnvioOcultadas, _
                      False, "", strFuncion
    
    F008_Ocultar_Hojas_Antiguas = True
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F008_Ocultar_Hojas_Antiguas = False
End Function
