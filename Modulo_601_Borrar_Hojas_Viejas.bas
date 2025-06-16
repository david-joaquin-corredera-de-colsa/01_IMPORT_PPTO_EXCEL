Attribute VB_Name = "Modulo_601_Borrar_Hojas_Viejas"
' =============================================================================
' FUNCION: Eliminar_Hojas_NoDeseadas
' FECHA Y HORA DE CREACION: 2025-06-13 12:09:30 UTC
' AUTOR: david-joaquin-corredera-de-colsa
' DESCRIPCION: Elimina hojas segun criterios especificos y ordena las pestañas
' PARAMETROS: vNumHojasIE_Target (Integer), vBorrarOtras_00 (Boolean),
'             vBorrarOtras_Import (Boolean), vBorrarOtras (Boolean)
' RETORNO: Boolean (True=exito, False=error)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function Eliminar_Hojas_NoDeseadas(ByVal vNumHojasIE_Target As Integer, _
                                      ByVal vBorrarOtras_00 As Boolean, _
                                      ByVal vBorrarOtras_Import As Boolean, _
                                      ByVal vBorrarOtras As Boolean) As Boolean

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables y optimizaciones de rendimiento
    ' 2. Validar parametros de entrada
    ' 3. Primera pasada: clasificar todas las hojas segun criterios
    ' 4. Segunda pasada: eliminar hojas segun los parametros booleanos
    ' 5. Tercera pasada: recuento y gestion de hojas Import_Envio_
    ' 6. Cuarta pasada: eliminar hojas Import_Envio_ excedentes si es necesario
    ' 7. Quinta pasada: ordenar pestañas segun criterios especificados
    ' 8. Restaurar configuraciones originales
    ' 9. Registrar resultado y retornar valor boolean

    On Error GoTo ErrorHandler
    
    ' Variables para control de errores y funcion
    Dim strFuncionActual As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResultado As Boolean
    
    ' Variables para optimizacion de rendimiento
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnEnableEventsOriginal As Boolean
    Dim lngCalculationOriginal As Long
    Dim blnDisplayAlertsOriginal As Boolean
    
    ' Variables para manejo de hojas
    Dim lngTotalHojas As Long
    Dim lngContadorHojas As Long
    Dim strNombreHoja As String
    Dim wsHojaActual As Worksheet
    
    ' Arrays para clasificacion de hojas
    Dim arrHojasProtegidas() As String
    Dim arrHojasImportEnvio() As String
    Dim arrHojasOtrasImport() As String
    Dim arrHojasDigitos() As String
    Dim arrHojasOtras() As String
    
    ' Contadores para arrays
    Dim lngNumHojasProtegidas As Long
    Dim lngNumHojasImportEnvio As Long
    Dim lngNumHojasOtrasImport As Long
    Dim lngNumHojasDigitos As Long
    Dim lngNumHojasOtras As Long
    
    ' Variables para gestion de Import_Envio_
    Dim lngNumeroHojasIE_Actuales As Long
    Dim lngHojasAEliminar As Long
    
    ' Variables auxiliares
    Dim i As Long, j As Long
    Dim strTempNombre As String
    Dim blnEsHojaProtegida As Boolean
    
    ' Inicializacion
    strFuncionActual = "Eliminar_Hojas_NoDeseadas"
    lngLineaError = 0
    blnResultado = False
    
    lngNumHojasProtegidas = 0
    lngNumHojasImportEnvio = 0
    lngNumHojasOtrasImport = 0
    lngNumHojasDigitos = 0
    lngNumHojasOtras = 0
    
    '--------------------------------------------------------------------------
    ' 1. Inicializar variables y optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    ' Registrar inicio de operacion
    Call fun801_LogMessage("Iniciando eliminacion y ordenamiento de hojas", False, "", strFuncionActual)
    
    ' Guardar configuraciones originales
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnEnableEventsOriginal = Application.EnableEvents
    lngCalculationOriginal = Application.Calculation
    blnDisplayAlertsOriginal = Application.DisplayAlerts
    
    ' Aplicar optimizaciones
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    '--------------------------------------------------------------------------
    ' 2. Validar parametros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 40
    
    ' Validar que vNumHojasIE_Target es un valor razonable
    If vNumHojasIE_Target < 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9001, strFuncionActual, _
            "Parametro vNumHojasIE_Target no puede ser negativo: " & vNumHojasIE_Target
    End If
    
    If vNumHojasIE_Target > 100 Then
        Call fun801_LogMessage("ADVERTENCIA: vNumHojasIE_Target muy alto: " & vNumHojasIE_Target, _
            False, "", strFuncionActual)
    End If
    
    ' Registrar parametros recibidos
    Call fun801_LogMessage("Parametros: Target=" & vNumHojasIE_Target & _
        ", Borrar00=" & vBorrarOtras_00 & ", BorrarImport=" & vBorrarOtras_Import & _
        ", BorrarOtras=" & vBorrarOtras, False, "", strFuncionActual)
    
    '--------------------------------------------------------------------------
    ' 3. Primera pasada: clasificar todas las hojas segun criterios
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    lngTotalHojas = ThisWorkbook.Worksheets.Count
    
    ' Redimensionar arrays con tamaño maximo posible
    ReDim arrHojasProtegidas(1 To lngTotalHojas)
    ReDim arrHojasImportEnvio(1 To lngTotalHojas)
    ReDim arrHojasOtrasImport(1 To lngTotalHojas)
    ReDim arrHojasDigitos(1 To lngTotalHojas)
    ReDim arrHojasOtras(1 To lngTotalHojas)
    
    ' Recorrer todas las hojas para clasificarlas
    For lngContadorHojas = 1 To lngTotalHojas
        strNombreHoja = ThisWorkbook.Worksheets(lngContadorHojas).Name
        
        ' Clasificar segun criterios especificados
        If fun801_EsHojaProtegidaDelSistema(strNombreHoja) Then
            ' Hojas protegidas del sistema
            lngNumHojasProtegidas = lngNumHojasProtegidas + 1
            arrHojasProtegidas(lngNumHojasProtegidas) = strNombreHoja
            
        ElseIf fun802_ComenzarConPrefijo(strNombreHoja, CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            ' Hojas Import_Envio_
            lngNumHojasImportEnvio = lngNumHojasImportEnvio + 1
            arrHojasImportEnvio(lngNumHojasImportEnvio) = strNombreHoja
            
        ElseIf fun802_ComenzarConPrefijo(strNombreHoja, CONST_PREFIJO_HOJA_IMPORTACION) Then
            ' Otras hojas Import_ (no Import_Envio_)
            lngNumHojasOtrasImport = lngNumHojasOtrasImport + 1
            arrHojasOtrasImport(lngNumHojasOtrasImport) = strNombreHoja
            
        ElseIf fun803_ComenzarConDigitosYGuion(strNombreHoja) Then
            ' Hojas que comienzan con dos digitos y guion bajo
            lngNumHojasDigitos = lngNumHojasDigitos + 1
            arrHojasDigitos(lngNumHojasDigitos) = strNombreHoja
            
        Else
            ' Todas las demas hojas
            lngNumHojasOtras = lngNumHojasOtras + 1
            arrHojasOtras(lngNumHojasOtras) = strNombreHoja
        End If
    Next lngContadorHojas
    
    ' Registrar clasificacion
    Call fun801_LogMessage("Clasificacion completada: Protegidas=" & lngNumHojasProtegidas & _
        ", ImportEnvio=" & lngNumHojasImportEnvio & ", OtrasImport=" & lngNumHojasOtrasImport & _
        ", Digitos=" & lngNumHojasDigitos & ", Otras=" & lngNumHojasOtras, _
        False, "", strFuncionActual)
    
    '--------------------------------------------------------------------------
    ' 4. Segunda pasada: eliminar hojas segun los parametros booleanos
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Eliminar hojas "Otras Import" si corresponde
    If vBorrarOtras_Import And lngNumHojasOtrasImport > 0 Then
        Call fun801_LogMessage("Eliminando " & lngNumHojasOtrasImport & " hojas Otras Import", _
            False, "", strFuncionActual)
        
        For i = 1 To lngNumHojasOtrasImport
            Call fun804_EliminarHojaSegura(arrHojasOtrasImport(i))
        Next i
    End If
    
    ' Eliminar hojas de digitos (no protegidas) si corresponde
    If vBorrarOtras_00 And lngNumHojasDigitos > 0 Then
        Call fun801_LogMessage("Eliminando hojas con digitos no protegidas", _
            False, "", strFuncionActual)
        
        For i = 1 To lngNumHojasDigitos
            ' Verificar que no sea una hoja protegida del sistema
            If Not fun801_EsHojaProtegidaDelSistema(arrHojasDigitos(i)) Then
                Call fun804_EliminarHojaSegura(arrHojasDigitos(i))
            End If
        Next i
    End If
    
    ' Eliminar otras hojas si corresponde
    If vBorrarOtras And lngNumHojasOtras > 0 Then
        Call fun801_LogMessage("Eliminando " & lngNumHojasOtras & " otras hojas", _
            False, "", strFuncionActual)
        
        For i = 1 To lngNumHojasOtras
            Call fun804_EliminarHojaSegura(arrHojasOtras(i))
        Next i
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Tercera pasada: recuento y gestion de hojas Import_Envio_
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Recalcular hojas Import_Envio_ actuales (pueden haber cambiado tras eliminaciones)
    Call fun805_RecalcularHojasImportEnvio(arrHojasImportEnvio, lngNumHojasImportEnvio)
    lngNumeroHojasIE_Actuales = lngNumHojasImportEnvio
    
    Call fun801_LogMessage("Hojas Import_Envio actuales: " & lngNumeroHojasIE_Actuales & _
        ", Target: " & vNumHojasIE_Target, False, "", strFuncionActual)
    
    '--------------------------------------------------------------------------
    ' 6. Cuarta pasada: eliminar hojas Import_Envio_ excedentes si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    If lngNumeroHojasIE_Actuales > vNumHojasIE_Target Then
        ' Hay que eliminar hojas excedentes
        lngHojasAEliminar = lngNumeroHojasIE_Actuales - vNumHojasIE_Target
        
        Call fun801_LogMessage("Eliminando " & lngHojasAEliminar & " hojas Import_Envio excedentes", _
            False, "", strFuncionActual)
        
        ' Ordenar hojas Import_Envio por orden lexicografico
        Call fun806_OrdenarArrayLexicografico(arrHojasImportEnvio, lngNumHojasImportEnvio)
        
        ' Eliminar las primeras (menores lexicograficamente)
        For i = 1 To lngHojasAEliminar
            Call fun804_EliminarHojaSegura(arrHojasImportEnvio(i))
        Next i
        
        ' Actualizar el array eliminando las hojas borradas
        Call fun807_ActualizarArrayTrasEliminacion(arrHojasImportEnvio, lngNumHojasImportEnvio, lngHojasAEliminar)
        lngNumHojasImportEnvio = lngNumHojasImportEnvio - lngHojasAEliminar
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Quinta pasada: ordenar pestañas segun criterios especificados
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    Call fun801_LogMessage("Iniciando ordenamiento de pestañas", False, "", strFuncionActual)
    Call fun808_OrdenarTodasLasPestanas
    
    '--------------------------------------------------------------------------
    ' 8. Operacion completada exitosamente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    blnResultado = True
    Call fun801_LogMessage("Eliminacion y ordenamiento completado exitosamente", _
        False, "", strFuncionActual)
    
    '--------------------------------------------------------------------------
    ' 9. Restaurar configuraciones originales
    '--------------------------------------------------------------------------
RestaurarConfiguracion:
    lngLineaError = 110
    
    ' Restaurar configuraciones originales
    Application.DisplayAlerts = blnDisplayAlertsOriginal
    Application.Calculation = lngCalculationOriginal
    Application.EnableEvents = blnEnableEventsOriginal
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    Eliminar_Hojas_NoDeseadas = blnResultado
    Exit Function
    
ErrorHandler:
    ' Construir mensaje de error detallado
    strMensajeError = "ERROR en " & strFuncionActual & vbCrLf & _
                      "Fecha: 2025-06-13 12:09:30 UTC" & vbCrLf & _
                      "Usuario: david-joaquin-corredera-de-colsa" & vbCrLf & _
                      "Linea aproximada: " & lngLineaError & vbCrLf & _
                      "Numero de Error: " & Err.Number & vbCrLf & _
                      "Descripcion: " & Err.Description & vbCrLf & _
                      "Parametros: Target=" & vNumHojasIE_Target & _
                      ", Borrar00=" & vBorrarOtras_00 & _
                      ", BorrarImport=" & vBorrarOtras_Import & _
                      ", BorrarOtras=" & vBorrarOtras
    
    ' Registrar error en log
    Call fun801_LogMessage(strMensajeError, True, "", strFuncionActual)
    
    ' Mostrar error al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncionActual
    
    ' Restaurar configuraciones en caso de error
    On Error Resume Next
    Application.DisplayAlerts = blnDisplayAlertsOriginal
    Application.Calculation = lngCalculationOriginal
    Application.EnableEvents = blnEnableEventsOriginal
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    On Error GoTo 0
    
    Eliminar_Hojas_NoDeseadas = False
End Function

' =============================================================================
' FUNCION AUXILIAR: fun801_EsHojaProtegidaDelSistema
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Verifica si una hoja esta en la lista de hojas protegidas del sistema
' PARAMETROS: strNombreHoja (String)
' RETORNO: Boolean (True=protegida, False=no protegida)
' =============================================================================
Public Function fun801_EsHojaProtegidaDelSistema(ByVal strNombreHoja As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim arrHojasProtegidas(1 To 7) As String
    Dim i As Integer
    
    ' Lista de hojas protegidas del sistema
    arrHojasProtegidas(1) = CONST_HOJA_EJECUTAR_PROCESOS
    arrHojasProtegidas(2) = CONST_HOJA_INVENTARIO
    arrHojasProtegidas(3) = CONST_HOJA_LOG
    arrHojasProtegidas(4) = CONST_HOJA_USERNAME
    arrHojasProtegidas(5) = CONST_HOJA_DELIMITADORES_ORIGINALES
    arrHojasProtegidas(6) = CONST_HOJA_REPORT_PL
    arrHojasProtegidas(7) = CONST_HOJA_REPORT_PL_AH
    
    fun801_EsHojaProtegidaDelSistema = False
    
    For i = 1 To 7
        If StrComp(strNombreHoja, arrHojasProtegidas(i), vbTextCompare) = 0 Then
            fun801_EsHojaProtegidaDelSistema = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    fun801_EsHojaProtegidaDelSistema = False
End Function

' =============================================================================
' FUNCION AUXILIAR: fun802_ComenzarConPrefijo
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Verifica si un nombre comienza con un prefijo especifico
' PARAMETROS: strNombre (String), strPrefijo (String)
' RETORNO: Boolean (True=comienza con prefijo, False=no comienza)
' =============================================================================
Public Function fun802_ComenzarConPrefijo(ByVal strNombre As String, ByVal strPrefijo As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    fun802_ComenzarConPrefijo = False
    
    If Len(strNombre) >= Len(strPrefijo) Then
        If StrComp(Left(strNombre, Len(strPrefijo)), strPrefijo, vbTextCompare) = 0 Then
            fun802_ComenzarConPrefijo = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    fun802_ComenzarConPrefijo = False
End Function

' =============================================================================
' FUNCION AUXILIAR: fun803_ComenzarConDigitosYGuion
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Verifica si un nombre comienza con dos digitos y guion bajo
' PARAMETROS: strNombre (String)
' RETORNO: Boolean (True=comienza con patron, False=no comienza)
' =============================================================================
Public Function fun803_ComenzarConDigitosYGuion(ByVal strNombre As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim strPrimerCaracter As String
    Dim strSegundoCaracter As String
    Dim strTercerCaracter As String
    
    fun803_ComenzarConDigitosYGuion = False
    
    ' Verificar que el nombre tenga al menos 3 caracteres
    If Len(strNombre) < 3 Then
        Exit Function
    End If
    
    ' Extraer los primeros tres caracteres
    strPrimerCaracter = Mid(strNombre, 1, 1)
    strSegundoCaracter = Mid(strNombre, 2, 1)
    strTercerCaracter = Mid(strNombre, 3, 1)
    
    ' Verificar patron: dos digitos seguidos de guion bajo
    If (strPrimerCaracter >= "0" And strPrimerCaracter <= "9") And _
       (strSegundoCaracter >= "0" And strSegundoCaracter <= "9") And _
       strTercerCaracter = Chr(95) Then
        fun803_ComenzarConDigitosYGuion = True
    End If
    
    Exit Function
    
ErrorHandler:
    fun803_ComenzarConDigitosYGuion = False
End Function

' =============================================================================
' SUB AUXILIAR: fun804_EliminarHojaSegura
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Elimina una hoja de forma segura con control de errores
' PARAMETROS: strNombreHoja (String)
' =============================================================================
Public Sub fun804_EliminarHojaSegura(ByVal strNombreHoja As String)
    
    On Error GoTo ErrorHandler
    
    Dim wsHoja As Worksheet
    Dim blnAlertasOriginales As Boolean
    
    ' Verificar que la hoja existe
    Set wsHoja = Nothing
    On Error Resume Next
    Set wsHoja = ThisWorkbook.Worksheets(strNombreHoja)
    On Error GoTo ErrorHandler
    
    If wsHoja Is Nothing Then
        Call fun801_LogMessage("ADVERTENCIA: Hoja no encontrada para eliminar: " & strNombreHoja, _
            False, "", "fun804_EliminarHojaSegura")
        Exit Sub
    End If
    
    ' Verificar que no es la unica hoja del libro
    If ThisWorkbook.Worksheets.Count <= 1 Then
        Call fun801_LogMessage("ADVERTENCIA: No se puede eliminar la unica hoja del libro: " & strNombreHoja, _
            False, "", "fun804_EliminarHojaSegura")
        Exit Sub
    End If
    
    ' Desactivar alertas temporalmente
    blnAlertasOriginales = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    ' Eliminar la hoja
    wsHoja.Delete
    
    ' Restaurar alertas
    Application.DisplayAlerts = blnAlertasOriginales
    
    Call fun801_LogMessage("Hoja eliminada: " & strNombreHoja, False, "", "fun804_EliminarHojaSegura")
    
    Exit Sub
    
ErrorHandler:
    ' Restaurar alertas en caso de error
    Application.DisplayAlerts = blnAlertasOriginales
    
    Call fun801_LogMessage("ERROR al eliminar hoja: " & strNombreHoja & " - " & Err.Description, _
        True, "", "fun804_EliminarHojaSegura")
End Sub

' =============================================================================
' SUB AUXILIAR: fun805_RecalcularHojasImportEnvio
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Recalcula las hojas Import_Envio existentes tras eliminaciones
' PARAMETROS: arrHojas (Array), lngNumHojas (Long por referencia)
' =============================================================================
Public Sub fun805_RecalcularHojasImportEnvio(ByRef arrHojas() As String, ByRef lngNumHojas As Long)
    
    On Error GoTo ErrorHandler
    
    Dim lngContador As Long
    Dim strNombreHoja As String
    Dim lngNuevoContador As Long
    
    lngNuevoContador = 0
    
    ' Recorrer todas las hojas del libro actual
    For lngContador = 1 To ThisWorkbook.Worksheets.Count
        strNombreHoja = ThisWorkbook.Worksheets(lngContador).Name
        
        If fun802_ComenzarConPrefijo(strNombreHoja, CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            lngNuevoContador = lngNuevoContador + 1
            If lngNuevoContador <= UBound(arrHojas) Then
                arrHojas(lngNuevoContador) = strNombreHoja
            End If
        End If
    Next lngContador
    
    lngNumHojas = lngNuevoContador
    
    Exit Sub
    
ErrorHandler:
    Call fun801_LogMessage("ERROR en fun805_RecalcularHojasImportEnvio: " & Err.Description, _
        True, "", "fun805_RecalcularHojasImportEnvio")
End Sub

' =============================================================================
' SUB AUXILIAR: fun806_OrdenarArrayLexicografico
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Ordena un array de strings lexicograficamente usando bubble sort
' PARAMETROS: arrTextos (Array), lngNumElementos (Long)
' =============================================================================
Public Sub fun806_OrdenarArrayLexicografico(ByRef arrTextos() As String, ByVal lngNumElementos As Long)
    
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim strTemporal As String
    
    ' Bubble sort compatible con Excel 97-365
    If lngNumElementos > 1 Then
        For i = 1 To lngNumElementos - 1
            For j = 1 To lngNumElementos - i
                If StrComp(arrTextos(j), arrTextos(j + 1), vbTextCompare) > 0 Then
                    strTemporal = arrTextos(j)
                    arrTextos(j) = arrTextos(j + 1)
                    arrTextos(j + 1) = strTemporal
                End If
            Next j
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    Call fun801_LogMessage("ERROR en fun806_OrdenarArrayLexicografico: " & Err.Description, _
        True, "", "fun806_OrdenarArrayLexicografico")
End Sub

' =============================================================================
' SUB AUXILIAR: fun807_ActualizarArrayTrasEliminacion
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Actualiza un array eliminando los primeros elementos
' PARAMETROS: arrTextos (Array), lngNumElementos (Long), lngElementosAEliminar (Long)
' =============================================================================
Public Sub fun807_ActualizarArrayTrasEliminacion(ByRef arrTextos() As String, _
                                                ByRef lngNumElementos As Long, _
                                                ByVal lngElementosAEliminar As Long)
    
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    ' Desplazar elementos hacia la izquierda
    For i = 1 To lngNumElementos - lngElementosAEliminar
        arrTextos(i) = arrTextos(i + lngElementosAEliminar)
    Next i
    
    ' Limpiar elementos al final
    For i = lngNumElementos - lngElementosAEliminar + 1 To lngNumElementos
        arrTextos(i) = ""
    Next i
    
    Exit Sub
    
ErrorHandler:
    Call fun801_LogMessage("ERROR en fun807_ActualizarArrayTrasEliminacion: " & Err.Description, _
        True, "", "fun807_ActualizarArrayTrasEliminacion")
End Sub

' =============================================================================
' SUB AUXILIAR: fun808_OrdenarTodasLasPestanas
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Ordena todas las pestañas segun los criterios especificados
' PARAMETROS: Ninguno
' =============================================================================
Public Sub fun808_OrdenarTodasLasPestanas()
    
    On Error GoTo ErrorHandler
    
    ' Arrays para almacenar hojas clasificadas por visibilidad y tipo
    Dim arrVisiblesDigitos() As String
    Dim arrVisiblesOtras() As String
    Dim arrOcultasDigitos() As String
    Dim arrOcultasOtras() As String
    
    ' Contadores para arrays
    Dim lngNumVisiblesDigitos As Long
    Dim lngNumVisiblesOtras As Long
    Dim lngNumOcultasDigitos As Long
    Dim lngNumOcultasOtras As Long
    
    ' Variables auxiliares
    Dim lngContador As Long
    Dim strNombreHoja As String
    Dim blnEsVisible As Boolean
    Dim lngTotalHojas As Long
    Dim lngPosicionActual As Long
    Dim i As Long
    
    lngTotalHojas = ThisWorkbook.Worksheets.Count
    lngPosicionActual = 1
    
    ' Inicializar contadores
    lngNumVisiblesDigitos = 0
    lngNumVisiblesOtras = 0
    lngNumOcultasDigitos = 0
    lngNumOcultasOtras = 0
    
    ' Redimensionar arrays
    ReDim arrVisiblesDigitos(1 To lngTotalHojas)
    ReDim arrVisiblesOtras(1 To lngTotalHojas)
    ReDim arrOcultasDigitos(1 To lngTotalHojas)
    ReDim arrOcultasOtras(1 To lngTotalHojas)
    
    ' Clasificar hojas por visibilidad y tipo
    For lngContador = 1 To lngTotalHojas
        strNombreHoja = ThisWorkbook.Worksheets(lngContador).Name
        blnEsVisible = (ThisWorkbook.Worksheets(strNombreHoja).Visible = xlSheetVisible)
        
        If blnEsVisible Then
            If fun803_ComenzarConDigitosYGuion(strNombreHoja) Then
                lngNumVisiblesDigitos = lngNumVisiblesDigitos + 1
                arrVisiblesDigitos(lngNumVisiblesDigitos) = strNombreHoja
            Else
                lngNumVisiblesOtras = lngNumVisiblesOtras + 1
                arrVisiblesOtras(lngNumVisiblesOtras) = strNombreHoja
            End If
        Else
            If fun803_ComenzarConDigitosYGuion(strNombreHoja) Then
                lngNumOcultasDigitos = lngNumOcultasDigitos + 1
                arrOcultasDigitos(lngNumOcultasDigitos) = strNombreHoja
            Else
                lngNumOcultasOtras = lngNumOcultasOtras + 1
                arrOcultasOtras(lngNumOcultasOtras) = strNombreHoja
            End If
        End If
    Next lngContador
    
    ' Ordenar cada array lexicograficamente
    Call fun806_OrdenarArrayLexicografico(arrVisiblesDigitos, lngNumVisiblesDigitos)
    Call fun806_OrdenarArrayLexicografico(arrVisiblesOtras, lngNumVisiblesOtras)
    Call fun806_OrdenarArrayLexicografico(arrOcultasDigitos, lngNumOcultasDigitos)
    Call fun806_OrdenarArrayLexicografico(arrOcultasOtras, lngNumOcultasOtras)
    
    ' Mover hojas segun el orden establecido
    ' 1. Visibles con digitos
    For i = 1 To lngNumVisiblesDigitos
        Call fun809_MoverHojaAPosicion(arrVisiblesDigitos(i), lngPosicionActual)
        lngPosicionActual = lngPosicionActual + 1
    Next i
    
    ' 2. Visibles otras
    For i = 1 To lngNumVisiblesOtras
        Call fun809_MoverHojaAPosicion(arrVisiblesOtras(i), lngPosicionActual)
        lngPosicionActual = lngPosicionActual + 1
    Next i
    
    ' 3. Ocultas con digitos
    For i = 1 To lngNumOcultasDigitos
        Call fun809_MoverHojaAPosicion(arrOcultasDigitos(i), lngPosicionActual)
        lngPosicionActual = lngPosicionActual + 1
    Next i
    
    ' 4. Ocultas otras
    For i = 1 To lngNumOcultasOtras
        Call fun809_MoverHojaAPosicion(arrOcultasOtras(i), lngPosicionActual)
        lngPosicionActual = lngPosicionActual + 1
    Next i
    
    Call fun801_LogMessage("Pestañas ordenadas: VisDigitos=" & lngNumVisiblesDigitos & _
        ", VisOtras=" & lngNumVisiblesOtras & ", OcultDigitos=" & lngNumOcultasDigitos & _
        ", OcultOtras=" & lngNumOcultasOtras, False, "", "fun808_OrdenarTodasLasPestanas")
    
    Exit Sub
    
ErrorHandler:
    Call fun801_LogMessage("ERROR en fun808_OrdenarTodasLasPestanas: " & Err.Description, _
        True, "", "fun808_OrdenarTodasLasPestanas")
End Sub

' =============================================================================
' SUB AUXILIAR: fun809_MoverHojaAPosicion
' FECHA: 2025-06-13 12:09:30 UTC
' DESCRIPCION: Mueve una hoja a una posicion especifica
' PARAMETROS: strNombreHoja (String), lngPosicion (Long)
' =============================================================================
Public Sub fun809_MoverHojaAPosicion(ByVal strNombreHoja As String, ByVal lngPosicion As Long)
    
    On Error GoTo ErrorHandler
    
    Dim wsHoja As Worksheet
    Dim lngTotalHojas As Long
    Dim lngPosicionActual As Long
    
    ' Verificar que la posicion es valida
    lngTotalHojas = ThisWorkbook.Worksheets.Count
    If lngPosicion < 1 Or lngPosicion > lngTotalHojas Then
        Exit Sub
    End If
    
    ' Verificar que la hoja existe
    Set wsHoja = Nothing
    On Error Resume Next
    Set wsHoja = ThisWorkbook.Worksheets(strNombreHoja)
    On Error GoTo ErrorHandler
    
    If wsHoja Is Nothing Then
        Exit Sub
    End If
    
    lngPosicionActual = wsHoja.Index
    
    ' Solo mover si la hoja no esta ya en la posicion correcta
    If lngPosicionActual <> lngPosicion Then
        If lngPosicion = 1 Then
            ' Mover al principio
            wsHoja.Move Before:=ThisWorkbook.Worksheets(1)
        Else
            ' Mover despues de la hoja en la posicion anterior
            wsHoja.Move After:=ThisWorkbook.Worksheets(lngPosicion - 1)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Error silencioso para no interrumpir el proceso
End Sub


Public Function Hoja_Tecnica_Visible() As Boolean
    
    '******************************************************************************
    ' FUNCIÓN: Hoja_Tecnica_Visible
    ' FECHA Y HORA DE CREACIÓN: 2025-06-15 07:13:34 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCIÓN:
    ' Determina si la hoja de trabajo actual debe estar visible según la configuración
    ' establecida en las constantes del sistema. Verifica si la hoja actual es una
    ' hoja técnica del sistema y consulta su configuración de visibilidad.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables de control de errores y optimización
    ' 2. Configuración de optimizaciones de rendimiento
    ' 3. Obtención del nombre de la hoja de trabajo actual usando ActiveSheet
    ' 4. Validación de que se obtuvo un nombre de hoja válido
    ' 5. Comparación case-insensitive con constantes de nombres de hojas técnicas
    ' 6. Búsqueda de la constante de visibilidad correspondiente
    ' 7. Evaluación del valor de visibilidad y determinación del resultado
    ' 8. Restauración de configuraciones de optimización
    ' 9. Retorno del resultado booleano
    ' 10. Manejo exhaustivo de errores con información detallada
    '
    ' PARÁMETROS: Ninguno
    ' RETORNA: Boolean - True si la hoja actual debe estar visible, False en caso contrario
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
    
    ' Variables para procesamiento
    Dim vThisWorkSheet As String
    Dim blnHojaEncontrada As Boolean
    Dim intConstanteVisibilidad As Integer
    
    ' Inicialización
    strFuncion = "Hoja_Tecnica_Visible"
    Hoja_Tecnica_Visible = False
    lngLineaError = 0
    blnHojaEncontrada = False
    intConstanteVisibilidad = xlSheetHidden
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables de control de errores y optimización
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    ' Almacenar configuraciones originales para restaurar después
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Configuración de optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 40
    
    ' Desactivar actualización de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar cálculo automático para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 3. Obtención del nombre de la hoja de trabajo actual usando ActiveSheet
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Verificar que hay una hoja activa disponible
    If ActiveSheet Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 8201, strFuncion, _
            "No hay hoja activa disponible en ThisWorkbook"
    End If
    
    ' Obtener nombre de la hoja actual (compatible Excel 97-365)
    vThisWorkSheet = ActiveSheet.Name
    
    '--------------------------------------------------------------------------
    ' 4. Validación de que se obtuvo un nombre de hoja válido
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Verificar que el nombre de hoja no esté vacío
    If Len(Trim(vThisWorkSheet)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 8202, strFuncion, _
            "El nombre de la hoja actual está vacío"
    End If
    
    ' Limpiar el nombre de espacios adicionales
    vThisWorkSheet = Trim(vThisWorkSheet)
    
    '--------------------------------------------------------------------------
    ' 5. Comparación case-insensitive con constantes de nombres de hojas técnicas
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Comparar con CONST_HOJA_EJECUTAR_PROCESOS usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_EJECUTAR_PROCESOS, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_EJECUTAR_PROCESOS_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    ' Comparar con CONST_HOJA_INVENTARIO usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_INVENTARIO, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_INVENTARIO_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    ' Comparar con CONST_HOJA_LOG usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_LOG, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_LOG_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    ' Comparar con CONST_HOJA_USERNAME usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_USERNAME, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_USERNAME_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    ' Comparar con CONST_HOJA_DELIMITADORES_ORIGINALES usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_DELIMITADORES_ORIGINALES, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    ' Comparar con CONST_HOJA_REPORT_PL usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_REPORT_PL, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_REPORT_PL_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    ' Comparar con CONST_HOJA_REPORT_PL_AH usando StrComp case-insensitive
    If StrComp(vThisWorkSheet, CONST_HOJA_REPORT_PL_AH, vbTextCompare) = 0 Then
        blnHojaEncontrada = True
        intConstanteVisibilidad = CONST_HOJA_REPORT_PL_AH_VISIBLE
        GoTo EvaluarVisibilidad
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Si no se encontró coincidencia, la hoja no es técnica del sistema
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    If Not blnHojaEncontrada Then
        ' La hoja actual no es una hoja técnica del sistema
        ' Registrar en log para información
        Call fun801_LogMessage("Hoja actual no es técnica del sistema: " & Chr(34) & vThisWorkSheet & Chr(34), _
            False, "", strFuncion)
        
        ' Retornar False ya que no es hoja técnica
        Hoja_Tecnica_Visible = False
        GoTo RestaurarConfiguracion
    End If

EvaluarVisibilidad:
    '--------------------------------------------------------------------------
    ' 7. Evaluación del valor de visibilidad y determinación del resultado
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    ' Verificar si el valor de la constante de visibilidad es xlSheetVisible
    If intConstanteVisibilidad = xlSheetVisible Then
        Hoja_Tecnica_Visible = True
        Call fun801_LogMessage("Hoja técnica debe estar VISIBLE: " & Chr(34) & vThisWorkSheet & Chr(34) & _
            " (Constante=" & intConstanteVisibilidad & ")", False, "", strFuncion)
    Else
        Hoja_Tecnica_Visible = False
        Call fun801_LogMessage("Hoja técnica debe estar OCULTA: " & Chr(34) & vThisWorkSheet & Chr(34) & _
            " (Constante=" & intConstanteVisibilidad & ")", False, "", strFuncion)
    End If

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 8. Restauración de configuraciones de optimización
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
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
    
    '--------------------------------------------------------------------------
    ' 9. Retorno del resultado booleano (ya establecido en secciones anteriores)
    '--------------------------------------------------------------------------
    lngLineaError = 110
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 10. Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja actual: " & Chr(34) & vThisWorkSheet & Chr(34) & vbCrLf & _
                      "Hoja encontrada: " & blnHojaEncontrada & vbCrLf & _
                      "Constante visibilidad: " & intConstanteVisibilidad & vbCrLf & _
                      "Fecha y Hora: " & Now() & vbCrLf & _
                      "Compatibilidad: Excel 97/2003/2007/365, OneDrive/SharePoint/Teams"
    
    ' Registrar error en log del sistema
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Para debugging en desarrollo
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
    
    ' Retornar False en caso de error
    Hoja_Tecnica_Visible = False
End Function
