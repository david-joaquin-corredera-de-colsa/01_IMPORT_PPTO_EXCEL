Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_06"
Option Explicit
Public Function Inventario_Actualizado_Si_No() As Boolean
    
    '******************************************************************************
    ' FUNCIÓN: Inventario_Actualizado_Si_No
    ' FECHA Y HORA DE CREACIÓN: 2025-01-15 14:30:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Compara el estado actual de las hojas del libro con la información almacenada
    ' en la hoja de inventario para determinar si el inventario está actualizado.
    ' Verifica tanto la existencia de hojas como su estado de visibilidad.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables y configuración de optimización
    ' 2. Recopilación de información actual de todas las hojas del libro
    ' 3. Lectura de información del inventario desde la hoja correspondiente
    ' 4. Comparación bidireccional entre arrays de hojas existentes e inventariadas
    ' 5. Validación de concordancia en nombres y estados de visibilidad
    ' 6. Generación de logging detallado de discrepancias encontradas
    ' 7. Restauración de configuración y retorno del resultado
    '
    ' PARÁMETROS: Ninguno
    ' RETORNA: Boolean - True si inventario actualizado, False si hay discrepancias
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimización
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas y datos
    Dim ws As Worksheet
    Dim wsInventario As Worksheet
    Dim lngTotalHojasLibro As Long
    Dim lngContadorHojas As Long
    Dim lngUltimaFilaInventario As Long
    Dim lngFilaActual As Long
    
    ' Arrays para almacenar información
    Dim vHojasExistentes() As Variant
    Dim vHojasInventariadas() As Variant
    Dim vNumeroHojasInventariadas As Integer
    Dim lngContadorInventario As Long
    
    ' Variables para comparación y validación
    Dim strNombreHoja As String
    Dim blnHojaVisible As Boolean
    Dim strValorColumnaVisible As String
    Dim blnEncontrado As Boolean
    Dim lngIndiceComparacion As Long
    
    ' Inicialización
    strFuncion = "Inventario_Actualizado_Si_No" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Inventario_Actualizado_Si_No"
    Inventario_Actualizado_Si_No = False
    lngLineaError = 0
    vNumeroHojasInventariadas = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables y configuración de optimización
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    Call fun801_LogMessage("Iniciando verificación de actualización del inventario", False, "", strFuncion)
    
    ' Guardar configuración original para restaurar después
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    ' Configurar optimización de rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 2. Recopilación de información actual de todas las hojas del libro
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Obtener número total de hojas en el libro
    lngTotalHojasLibro = ThisWorkbook.Worksheets.Count
    
    If lngTotalHojasLibro = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9001, strFuncion, _
            "No hay hojas en el libro de trabajo"
    End If
    
    ' Dimensionar array para hojas existentes (2 dimensiones)
    ReDim vHojasExistentes(1 To lngTotalHojasLibro, 1 To 2)
    
    Call fun801_LogMessage("Recopilando información de " & lngTotalHojasLibro & " hojas existentes", _
        False, "", strFuncion)
    
    ' Recorrer todas las hojas del libro y recopilar información
    For lngContadorHojas = 1 To lngTotalHojasLibro
        lngLineaError = 70 + lngContadorHojas
        
        Set ws = ThisWorkbook.Worksheets(lngContadorHojas)
        
        ' Almacenar nombre de la hoja (dimensión 1)
        vHojasExistentes(lngContadorHojas, 1) = ws.Name
        
        ' Almacenar estado de visibilidad (dimensión 2)
        ' True si visible, False si oculta
        vHojasExistentes(lngContadorHojas, 2) = (ws.Visible = xlSheetVisible)
        
        Call fun801_LogMessage("Hoja " & lngContadorHojas & ": " & Chr(34) & ws.Name & Chr(34) & _
            " - Visible: " & CStr(vHojasExistentes(lngContadorHojas, 2)), False, "", strFuncion)
    Next lngContadorHojas
    
    '--------------------------------------------------------------------------
    ' 3. Lectura de información del inventario desde la hoja correspondiente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Verificar existencia de hoja de inventario
    If Not fun802_SheetExists(CONST_HOJA_INVENTARIO) Then
        Err.Raise ERROR_BASE_IMPORT + 9002, strFuncion, _
            "La hoja de inventario no existe: " & CONST_HOJA_INVENTARIO
    End If
    
    Set wsInventario = ThisWorkbook.Worksheets(CONST_HOJA_INVENTARIO)
    
    ' Encontrar última fila con datos en la columna de nombres
    lngUltimaFilaInventario = wsInventario.Cells(wsInventario.Rows.Count, CONST_INVENTARIO_COLUMNA_NOMBRE).End(xlUp).Row
    
    Call fun801_LogMessage("Última fila con datos en inventario: " & lngUltimaFilaInventario, _
        False, "", strFuncion)
    
    ' Verificar que hay datos después de los headers
    If lngUltimaFilaInventario <= CONST_INVENTARIO_FILA_HEADERS Then
        Call fun801_LogMessage("WARNING: No hay datos en el inventario después de la fila de headers", _
            True, "", strFuncion)
        GoTo RestaurarConfiguracion ' Considerar como no actualizado
    End If
    
    '--------------------------------------------------------------------------
    ' 3.1. Contar hojas inventariadas (con datos válidos)
    '--------------------------------------------------------------------------
    lngLineaError = 110
    
    vNumeroHojasInventariadas = 0
    
    ' Recorrer filas del inventario para contar las que tienen nombre de hoja
    For lngFilaActual = CONST_INVENTARIO_FILA_HEADERS + 1 To lngUltimaFilaInventario
        strNombreHoja = Trim(CStr(wsInventario.Cells(lngFilaActual, CONST_INVENTARIO_COLUMNA_NOMBRE).Value))
        
        If Len(strNombreHoja) > 0 Then
            vNumeroHojasInventariadas = vNumeroHojasInventariadas + 1
        End If
    Next lngFilaActual
    
    Call fun801_LogMessage("Número de hojas inventariadas con datos válidos: " & vNumeroHojasInventariadas, _
        False, "", strFuncion)
    
    If vNumeroHojasInventariadas = 0 Then
        Call fun801_LogMessage("WARNING: No hay hojas inventariadas con datos válidos", _
            True, "", strFuncion)
        GoTo RestaurarConfiguracion ' Considerar como no actualizado
    End If
    
    '--------------------------------------------------------------------------
    ' 3.2. Llenar array de hojas inventariadas
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
    ' Dimensionar array para hojas inventariadas
    ReDim vHojasInventariadas(1 To vNumeroHojasInventariadas, 1 To 2)
    
    lngContadorInventario = 0
    
    ' Llenar array con datos del inventario
    For lngFilaActual = CONST_INVENTARIO_FILA_HEADERS + 1 To lngUltimaFilaInventario
        lngLineaError = 130 + lngFilaActual
        
        strNombreHoja = Trim(CStr(wsInventario.Cells(lngFilaActual, CONST_INVENTARIO_COLUMNA_NOMBRE).Value))
        
        If Len(strNombreHoja) > 0 Then
            lngContadorInventario = lngContadorInventario + 1
            
            ' Almacenar nombre de hoja (dimensión 1)
            vHojasInventariadas(lngContadorInventario, 1) = strNombreHoja
            
            ' Obtener y transformar valor de visibilidad (dimensión 2)
            strValorColumnaVisible = Trim(CStr(wsInventario.Cells(lngFilaActual, CONST_INVENTARIO_COLUMNA_VISIBLE).Value))
            
            ' Transformar según especificaciones:
            ' "OCULTA" -> False (hoja oculta)
            ' ">> visible <<" -> True (hoja visible)
            If StrComp(strValorColumnaVisible, "OCULTA", vbTextCompare) = 0 Then
                vHojasInventariadas(lngContadorInventario, 2) = False
            ElseIf StrComp(strValorColumnaVisible, ">> visible <<", vbTextCompare) = 0 Then
                vHojasInventariadas(lngContadorInventario, 2) = True
            Else
                ' Valor no reconocido, asumir visible por defecto y registrar warning
                vHojasInventariadas(lngContadorInventario, 2) = True
                Call fun801_LogMessage("WARNING: Valor de visibilidad no reconocido para hoja " & Chr(34) & _
                    strNombreHoja & Chr(34) & ": " & Chr(34) & strValorColumnaVisible & Chr(34) & _
                    ". Asumiendo visible.", True, "", strFuncion)
            End If
            
            Call fun801_LogMessage("Inventario " & lngContadorInventario & ": " & Chr(34) & strNombreHoja & _
                Chr(34) & " - Visible: " & CStr(vHojasInventariadas(lngContadorInventario, 2)), _
                False, "", strFuncion)
        End If
    Next lngFilaActual
    
    '--------------------------------------------------------------------------
    ' 4. Comparación bidireccional entre arrays
    '--------------------------------------------------------------------------
    lngLineaError = 200
    
    Call fun801_LogMessage("Iniciando comparación bidireccional de arrays", False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 4.1. Verificar que cada hoja existente está en el inventario
    '--------------------------------------------------------------------------
    lngLineaError = 210
    
    For lngContadorHojas = 1 To lngTotalHojasLibro
        lngLineaError = 220 + lngContadorHojas
        
        strNombreHoja = CStr(vHojasExistentes(lngContadorHojas, 1))
        blnHojaVisible = CBool(vHojasExistentes(lngContadorHojas, 2))
        blnEncontrado = False
        
        ' Buscar la hoja actual en el inventario
        For lngIndiceComparacion = 1 To vNumeroHojasInventariadas
            If StrComp(CStr(vHojasInventariadas(lngIndiceComparacion, 1)), strNombreHoja, vbTextCompare) = 0 Then
                blnEncontrado = True
                
                ' Comparar estado de visibilidad
                If CBool(vHojasInventariadas(lngIndiceComparacion, 2)) <> blnHojaVisible Then
                    Call fun801_LogMessage("DISCREPANCIA: Hoja " & Chr(34) & strNombreHoja & Chr(34) & _
                        " - Estado actual: " & CStr(blnHojaVisible) & _
                        ", Estado en inventario: " & CStr(vHojasInventariadas(lngIndiceComparacion, 2)), _
                        True, "", strFuncion)
                    GoTo RestaurarConfiguracion ' Retornar False
                End If
                Exit For
            End If
        Next lngIndiceComparacion
        
        ' Si la hoja no se encontró en el inventario
        If Not blnEncontrado Then
            Call fun801_LogMessage("DISCREPANCIA: Hoja existente " & Chr(34) & strNombreHoja & _
                Chr(34) & " no encontrada en el inventario", True, "", strFuncion)
            GoTo RestaurarConfiguracion ' Retornar False
        End If
    Next lngContadorHojas
    
    '--------------------------------------------------------------------------
    ' 4.2. Verificar que cada hoja inventariada existe realmente
    '--------------------------------------------------------------------------
    lngLineaError = 250
    
    For lngContadorInventario = 1 To vNumeroHojasInventariadas
        lngLineaError = 260 + lngContadorInventario
        
        strNombreHoja = CStr(vHojasInventariadas(lngContadorInventario, 1))
        blnHojaVisible = CBool(vHojasInventariadas(lngContadorInventario, 2))
        blnEncontrado = False
        
        ' Buscar la hoja inventariada en las hojas existentes
        For lngIndiceComparacion = 1 To lngTotalHojasLibro
            If StrComp(CStr(vHojasExistentes(lngIndiceComparacion, 1)), strNombreHoja, vbTextCompare) = 0 Then
                blnEncontrado = True
                
                ' Comparar estado de visibilidad
                If CBool(vHojasExistentes(lngIndiceComparacion, 2)) <> blnHojaVisible Then
                    Call fun801_LogMessage("DISCREPANCIA: Hoja inventariada " & Chr(34) & strNombreHoja & _
                        Chr(34) & " - Estado en inventario: " & CStr(blnHojaVisible) & _
                        ", Estado actual: " & CStr(vHojasExistentes(lngIndiceComparacion, 2)), _
                        True, "", strFuncion)
                    GoTo RestaurarConfiguracion ' Retornar False
                End If
                Exit For
            End If
        Next lngIndiceComparacion
        
        ' Si la hoja inventariada no existe realmente
        If Not blnEncontrado Then
            Call fun801_LogMessage("DISCREPANCIA: Hoja inventariada " & Chr(34) & strNombreHoja & _
                Chr(34) & " no existe en el libro actual", True, "", strFuncion)
            GoTo RestaurarConfiguracion ' Retornar False
        End If
    Next lngContadorInventario
    
    '--------------------------------------------------------------------------
    ' 5. Si llegamos aquí, el inventario está actualizado
    '--------------------------------------------------------------------------
    lngLineaError = 300
    
    Call fun801_LogMessage("ÉXITO: El inventario está completamente actualizado. " & _
        "Hojas existentes: " & lngTotalHojasLibro & ", Hojas inventariadas: " & vNumeroHojasInventariadas, _
        False, "", strFuncion)
    
    Inventario_Actualizado_Si_No = True

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 6. Restauración de configuración y limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 350
    
    ' Restaurar configuración original
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set ws = Nothing
    Set wsInventario = Nothing
    
    Call fun801_LogMessage("Verificación de inventario completada. Resultado: " & _
        CStr(Inventario_Actualizado_Si_No), False, "", strFuncion)
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 7. Manejo exhaustivo de errores
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hojas en libro: " & lngTotalHojasLibro & vbCrLf & _
                      "Hojas inventariadas: " & vNumeroHojasInventariadas & vbCrLf & _
                      "Hoja actual procesando: " & strNombreHoja & vbCrLf & _
                      "Fecha y Hora: " & Now()
    
    ' Registrar error en log
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Mostrar error al usuario (opcional)
    MsgBox strMensajeError, vbCritical, "Error en Verificación de Inventario"
    
    ' Restaurar configuración en caso de error
    On Error Resume Next
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias
    Set ws = Nothing
    Set wsInventario = Nothing
    
    ' Retornar False en caso de error
    Inventario_Actualizado_Si_No = False
End Function
' =============================================================================
' FUNCION: Ordenar_Hojas
' FECHA: 2025-06-13 08:28:44 UTC
' DESCRIPCION: Ordena las pestañas del libro con prioridad por visibilidad y formato de nombre
' PARAMETROS: Ninguno
' RETORNO: Boolean (True=éxito, False=error)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function Ordenar_Hojas() As Boolean

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Optimizar configuración de Excel para mejor rendimiento
    ' 2. Recopilar información de todas las hojas del libro
    ' 3. Separar hojas visibles y ocultas en arrays independientes
    ' 4. Categorizar cada grupo por patrón de nombre (con/sin prefijo numérico)
    ' 5. Ordenar lexicográficamente cada subcategoría por separado
    ' 6. Reorganizar las hojas según el orden establecido
    ' 7. Restaurar configuración original de Excel
    ' 8. Retornar resultado de la operación

    On Error GoTo ErrorHandler
    
    Dim vResultado As Boolean
    Dim vLineaError As Integer
    Dim vTotalHojas As Integer
    Dim vContadorHojas As Integer
    Dim vNombreHoja As String
    Dim vEsVisible As Boolean
    
    ' Arrays para almacenar hojas visibles categorizadas
    Dim vHojasVisiblesConPrefijo() As String
    Dim vHojasVisiblesSinPrefijo() As String
    Dim vNumVisiblesConPrefijo As Integer
    Dim vNumVisiblesSinPrefijo As Integer
    
    ' Arrays para almacenar hojas ocultas categorizadas
    Dim vHojasOcultasConPrefijo() As String
    Dim vHojasOcultasSinPrefijo() As String
    Dim vNumOcultasConPrefijo As Integer
    Dim vNumOcultasSinPrefijo As Integer
    
    ' Variables para ordenamiento y control
    Dim i As Integer, j As Integer
    Dim vTempNombre As String
    Dim vPosicionActual As Integer
    
    ' Variables para optimización (inicialización correcta)
    Dim vCalculationOriginal As Integer
    Dim vScreenUpdatingOriginal As Boolean
    Dim vEnableEventsOriginal As Boolean
    
    ' Variables para manejo de alertas
    Dim vDisplayAlertsOriginal As Boolean
    
    ' Inicialización de variables
    vResultado = False
    vLineaError = 10
    vNumVisiblesConPrefijo = 0
    vNumVisiblesSinPrefijo = 0
    vNumOcultasConPrefijo = 0
    vNumOcultasSinPrefijo = 0
    vPosicionActual = 1
    
    ' Paso 1: Optimizar configuración de Excel para mejor rendimiento
    vLineaError = 20
    
    ' Guardar configuración original
    vCalculationOriginal = Application.Calculation
    vScreenUpdatingOriginal = Application.ScreenUpdating
    vEnableEventsOriginal = Application.EnableEvents
    vDisplayAlertsOriginal = Application.DisplayAlerts
    
    ' Aplicar optimizaciones
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Registrar inicio de operación en log (con control de errores)
    On Error Resume Next
    Call fun801_LogMessage("Iniciando ordenamiento avanzado de hojas", False, "", "Ordenar_Hojas")
    On Error GoTo ErrorHandler
    
    ' Paso 2: Recopilar información de todas las hojas del libro
    vLineaError = 30
    vTotalHojas = ThisWorkbook.Worksheets.Count
    
    ' Validar que hay hojas para procesar
    If vTotalHojas <= 1 Then
        vResultado = True ' No hay nada que ordenar, pero no es error
        GoTo RestaurarConfiguracion
    End If
    
    ' Redimensionar arrays con tamaño máximo posible
    ReDim vHojasVisiblesConPrefijo(1 To vTotalHojas)
    ReDim vHojasVisiblesSinPrefijo(1 To vTotalHojas)
    ReDim vHojasOcultasConPrefijo(1 To vTotalHojas)
    ReDim vHojasOcultasSinPrefijo(1 To vTotalHojas)
    
    ' Paso 3: Separar hojas visibles y ocultas en arrays independientes
    ' Paso 4: Categorizar cada grupo por patrón de nombre
    vLineaError = 40
    For vContadorHojas = 1 To vTotalHojas
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        vEsVisible = (ThisWorkbook.Worksheets(vNombreHoja).Visible = xlSheetVisible)
        
        If vEsVisible Then
            ' Hoja visible: categorizar por patrón de nombre
            If fun801_TienePrefijoNumerico(vNombreHoja) Then
                vNumVisiblesConPrefijo = vNumVisiblesConPrefijo + 1
                vHojasVisiblesConPrefijo(vNumVisiblesConPrefijo) = vNombreHoja
            Else
                vNumVisiblesSinPrefijo = vNumVisiblesSinPrefijo + 1
                vHojasVisiblesSinPrefijo(vNumVisiblesSinPrefijo) = vNombreHoja
            End If
        Else
            ' Hoja oculta: categorizar por patrón de nombre
            If fun801_TienePrefijoNumerico(vNombreHoja) Then
                vNumOcultasConPrefijo = vNumOcultasConPrefijo + 1
                vHojasOcultasConPrefijo(vNumOcultasConPrefijo) = vNombreHoja
            Else
                vNumOcultasSinPrefijo = vNumOcultasSinPrefijo + 1
                vHojasOcultasSinPrefijo(vNumOcultasSinPrefijo) = vNombreHoja
            End If
        End If
    Next vContadorHojas
    
    ' Paso 5: Ordenar lexicográficamente cada subcategoría por separado
    vLineaError = 50
    
    ' Ordenar hojas visibles con prefijo numérico
    If vNumVisiblesConPrefijo > 1 Then
        For i = 1 To vNumVisiblesConPrefijo - 1
            For j = 1 To vNumVisiblesConPrefijo - i
                If StrComp(vHojasVisiblesConPrefijo(j), vHojasVisiblesConPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasVisiblesConPrefijo(j)
                    vHojasVisiblesConPrefijo(j) = vHojasVisiblesConPrefijo(j + 1)
                    vHojasVisiblesConPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Ordenar hojas visibles sin prefijo numérico
    If vNumVisiblesSinPrefijo > 1 Then
        For i = 1 To vNumVisiblesSinPrefijo - 1
            For j = 1 To vNumVisiblesSinPrefijo - i
                If StrComp(vHojasVisiblesSinPrefijo(j), vHojasVisiblesSinPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasVisiblesSinPrefijo(j)
                    vHojasVisiblesSinPrefijo(j) = vHojasVisiblesSinPrefijo(j + 1)
                    vHojasVisiblesSinPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Ordenar hojas ocultas con prefijo numérico
    If vNumOcultasConPrefijo > 1 Then
        For i = 1 To vNumOcultasConPrefijo - 1
            For j = 1 To vNumOcultasConPrefijo - i
                If StrComp(vHojasOcultasConPrefijo(j), vHojasOcultasConPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasOcultasConPrefijo(j)
                    vHojasOcultasConPrefijo(j) = vHojasOcultasConPrefijo(j + 1)
                    vHojasOcultasConPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Ordenar hojas ocultas sin prefijo numérico
    If vNumOcultasSinPrefijo > 1 Then
        For i = 1 To vNumOcultasSinPrefijo - 1
            For j = 1 To vNumOcultasSinPrefijo - i
                If StrComp(vHojasOcultasSinPrefijo(j), vHojasOcultasSinPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasOcultasSinPrefijo(j)
                    vHojasOcultasSinPrefijo(j) = vHojasOcultasSinPrefijo(j + 1)
                    vHojasOcultasSinPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Paso 6: Reorganizar las hojas según el orden establecido
    vLineaError = 60
    
    ' 6.1: Primero las hojas visibles con prefijo numérico
    For i = 1 To vNumVisiblesConPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasVisiblesConPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    ' 6.2: Después las hojas visibles sin prefijo numérico
    For i = 1 To vNumVisiblesSinPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasVisiblesSinPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    ' 6.3: Después las hojas ocultas con prefijo numérico
    For i = 1 To vNumOcultasConPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasOcultasConPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    ' 6.4: Finalmente las hojas ocultas sin prefijo numérico
    For i = 1 To vNumOcultasSinPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasOcultasSinPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    vResultado = True
    
RestaurarConfiguracion:
    ' Paso 7: Restaurar configuración original de Excel
    vLineaError = 70
    On Error Resume Next
    Application.DisplayAlerts = vDisplayAlertsOriginal
    Application.EnableEvents = vEnableEventsOriginal
    Application.ScreenUpdating = vScreenUpdatingOriginal
    Application.Calculation = vCalculationOriginal
    On Error GoTo 0
    
    ' Registrar finalización en log (con control de errores)
    If vResultado Then
        On Error Resume Next
        Call fun801_LogMessage("Ordenamiento de hojas completado exitosamente. Total procesadas: " & _
            CStr(vTotalHojas) & ", Visibles con prefijo: " & CStr(vNumVisiblesConPrefijo) & _
            ", Visibles sin prefijo: " & CStr(vNumVisiblesSinPrefijo) & _
            ", Ocultas con prefijo: " & CStr(vNumOcultasConPrefijo) & _
            ", Ocultas sin prefijo: " & CStr(vNumOcultasSinPrefijo), False, "", "Ordenar_Hojas")
        On Error GoTo 0
    End If
    
    ' Paso 8: Retornar resultado de la operación
    Ordenar_Hojas = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en Ordenar_Hojas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Usuario: david-joaquin-corredera-de-colsa" & vbCrLf & _
                   "Fecha y Hora: 2025-06-13 08:28:44 UTC"
    
    ' Restaurar configuración en caso de error
    On Error Resume Next
    Application.DisplayAlerts = vDisplayAlertsOriginal
    Application.EnableEvents = vEnableEventsOriginal
    Application.ScreenUpdating = vScreenUpdatingOriginal
    Application.Calculation = vCalculationOriginal
    On Error GoTo 0
    
    ' Registrar error en log
    On Error Resume Next
    Call fun801_LogMessage(vMensajeError, True, "", "Ordenar_Hojas")
    On Error GoTo 0
    
    MsgBox vMensajeError, vbCritical, "Error Ordenar_Hojas"
    
    Ordenar_Hojas = False
    
End Function

' =============================================================================
' FUNCION AUXILIAR: fun801_TienePrefijoNumerico
' FECHA: 2025-06-13 08:28:44 UTC
' DESCRIPCION: Verifica si el nombre de hoja tiene prefijo con formato "##_"
' PARAMETROS: vNombreHoja (String)
' RETORNO: Boolean (True si tiene prefijo numérico, False si no)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function fun801_TienePrefijoNumerico(vNombreHoja As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim vPrimerCaracter As String
    Dim vSegundoCaracter As String
    Dim vTercerCaracter As String
    
    ' Inicialización
    fun801_TienePrefijoNumerico = False
    
    ' Verificar que el nombre tenga al menos 3 caracteres
    If Len(vNombreHoja) < 3 Then
        Exit Function
    End If
    
    ' Extraer los primeros tres caracteres
    vPrimerCaracter = Mid(vNombreHoja, 1, 1)
    vSegundoCaracter = Mid(vNombreHoja, 2, 1)
    vTercerCaracter = Mid(vNombreHoja, 3, 1)
    
    ' Verificar patrón: dos dígitos seguidos de guión bajo
    ' Usar verificación manual para compatibilidad con Excel 97
    If (vPrimerCaracter >= "0" And vPrimerCaracter <= "9") And _
       (vSegundoCaracter >= "0" And vSegundoCaracter <= "9") And _
       vTercerCaracter = Chr(95) Then
        fun801_TienePrefijoNumerico = True
    End If
    
    Exit Function
    
ErrorHandler:
    fun801_TienePrefijoNumerico = False
    
End Function

' =============================================================================
' SUB AUXILIAR: fun803_Mover_Hoja_A_Posicion_Segura
' FECHA: 2025-06-13 08:28:44 UTC
' DESCRIPCION: Mueve una hoja a una posición específica con control de errores
' PARAMETROS: vNombreHoja (String), vPosicion (Integer)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Sub fun803_Mover_Hoja_A_Posicion_Segura(vNombreHoja As String, vPosicion As Integer)
    
    On Error GoTo ErrorHandler
    
    Dim vHoja As Worksheet
    Dim vTotalHojas As Integer
    Dim vPosicionActualHoja As Integer
    Dim vHojaReferencia As Worksheet
    
    ' Verificar que la posición es válida
    vTotalHojas = ThisWorkbook.Worksheets.Count
    If vPosicion < 1 Or vPosicion > vTotalHojas Then
        Exit Sub
    End If
    
    ' Verificar que la hoja existe
    Set vHoja = Nothing
    On Error Resume Next
    Set vHoja = ThisWorkbook.Worksheets(vNombreHoja)
    On Error GoTo ErrorHandler
    
    If vHoja Is Nothing Then
        Exit Sub
    End If
    
    vPosicionActualHoja = vHoja.Index
    
    ' Solo mover si la hoja no está ya en la posición correcta
    If vPosicionActualHoja <> vPosicion Then
        ' Mover la hoja a la posición especificada
        If vPosicion = 1 Then
            ' Si es la primera posición, mover antes de la primera hoja
            vHoja.Move Before:=ThisWorkbook.Worksheets(1)
        Else
            ' Obtener referencia a la hoja en la posición objetivo
            Set vHojaReferencia = Nothing
            On Error Resume Next
            
            If vPosicionActualHoja < vPosicion Then
                ' La hoja está antes de su destino
                Set vHojaReferencia = ThisWorkbook.Worksheets(vPosicion - 1)
                If Not vHojaReferencia Is Nothing Then
                    vHoja.Move After:=vHojaReferencia
                End If
            Else
                ' La hoja está después de su destino
                Set vHojaReferencia = ThisWorkbook.Worksheets(vPosicion)
                If Not vHojaReferencia Is Nothing Then
                    vHoja.Move Before:=vHojaReferencia
                End If
            End If
            
            On Error GoTo ErrorHandler
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Registrar error específico en log si es posible
    On Error Resume Next
    Call fun801_LogMessage("Error al mover hoja " & Chr(34) & vNombreHoja & Chr(34) & _
        " a posición " & CStr(vPosicion) & ": " & Err.Description, True, "", "fun803_Mover_Hoja_A_Posicion_Segura")
    On Error GoTo 0
    
End Sub

