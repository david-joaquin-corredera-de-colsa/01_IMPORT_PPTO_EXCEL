Attribute VB_Name = "Modulo_030_Forzar_Delim"
' =============================================================================
' M�DULO: F004_Forzar_Delimitadores_en_Excel
' FECHA CREACI�N: 2025-05-26 15:11:21 UTC
' AUTOR: Sistema Automatizado
' VERSI�N: 1.0
' COMPATIBILIDAD: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

' Variables globales para delimitadores
Public vDelimitadorDecimal_HFM As String
Public vDelimitadorMiles_HFM As String

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
Public Function F004_Forzar_Delimitadores_en_Excel() As Boolean

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

' =============================================================================
' FUNCI�N: fun801_InicializarVariablesGlobales
' PROP�SITO: Inicializa las variables globales con valores por defecto
' FECHA: 2025-05-26 15:11:21 UTC
' =============================================================================
Private Sub fun801_InicializarVariablesGlobales()
    On Error GoTo ErrorHandler_fun801
    
    ' Inicializar delimitador decimal si est� vac�o
    If vDelimitadorDecimal_HFM = "" Or vDelimitadorDecimal_HFM = vbNullString Then
        vDelimitadorDecimal_HFM = "."
    End If
    
    ' Inicializar delimitador de miles si est� vac�o
    If vDelimitadorMiles_HFM = "" Or vDelimitadorMiles_HFM = vbNullString Then
        vDelimitadorMiles_HFM = ","
    End If
    
    Exit Sub

ErrorHandler_fun801:
    ' En caso de error, usar valores por defecto
    vDelimitadorDecimal_HFM = "."
    vDelimitadorMiles_HFM = ","
End Sub

' =============================================================================
' FUNCI�N: fun802_VerificarCompatibilidad
' PROP�SITO: Verifica compatibilidad con diferentes versiones de Excel
' FECHA: 2025-05-26 15:11:21 UTC
' RETORNA: Boolean (True = compatible, False = no compatible)
' =============================================================================
Private Function fun802_VerificarCompatibilidad() As Boolean
    On Error GoTo ErrorHandler_fun802
    
    Dim strVersionExcel As String
    Dim dblVersionNumero As Double
    
    ' Obtener versi�n de Excel
    strVersionExcel = Application.Version
    dblVersionNumero = CDbl(strVersionExcel)
    
    ' Verificar compatibilidad (Excel 97 = 8.0, 2003 = 11.0, 365 = 16.0+)
    If dblVersionNumero >= 8# Then
        fun802_VerificarCompatibilidad = True
    Else
        fun802_VerificarCompatibilidad = False
    End If
    
    Exit Function

ErrorHandler_fun802:
    ' En caso de error, asumir compatibilidad
    fun802_VerificarCompatibilidad = True
End Function

' =============================================================================
' FUNCI�N: fun803_ObtenerConfiguracionActual
' PROP�SITO: Obtiene la configuraci�n actual de delimitadores
' FECHA: 2025-05-26 15:11:21 UTC
' =============================================================================
Private Sub fun803_ObtenerConfiguracionActual(ByRef strDecimalAnterior As String, ByRef strMilesAnterior As String)
    On Error GoTo ErrorHandler_fun803
    
    ' Obtener delimitador decimal actual
    strDecimalAnterior = Application.International(xlDecimalSeparator)
    
    ' Obtener delimitador de miles actual
    strMilesAnterior = Application.International(xlThousandsSeparator)
    
    Exit Sub

ErrorHandler_fun803:
    ' En caso de error, usar valores por defecto
    strDecimalAnterior = "."
    strMilesAnterior = ","
End Sub

' =============================================================================
' FUNCI�N: fun804_AplicarNuevosDelimitadores
' PROP�SITO: Aplica los nuevos delimitadores al sistema
' FECHA: 2025-05-26 15:11:21 UTC
' RETORNA: Boolean (True = �xito, False = error)
' =============================================================================
Private Function fun804_AplicarNuevosDelimitadores() As Boolean
    On Error GoTo ErrorHandler_fun804
    
    ' Aplicar nuevo delimitador decimal
    Application.DecimalSeparator = vDelimitadorDecimal_HFM
    
    ' Aplicar nuevo delimitador de miles
    Application.ThousandsSeparator = vDelimitadorMiles_HFM
    
    ' Forzar que Excel use los delimitadores del sistema
    Application.UseSystemSeparators = False
    
    ' Actualizar la pantalla
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    fun804_AplicarNuevosDelimitadores = True
    Exit Function

ErrorHandler_fun804:
    fun804_AplicarNuevosDelimitadores = False
End Function

' =============================================================================
' FUNCI�N: fun805_VerificarAplicacionDelimitadores
' PROP�SITO: Verifica que los delimitadores se aplicaron correctamente
' FECHA: 2025-05-26 15:11:21 UTC
' RETORNA: Boolean (True = aplicados correctamente, False = error)
' =============================================================================
Private Function fun805_VerificarAplicacionDelimitadores() As Boolean
    On Error GoTo ErrorHandler_fun805
    
    Dim strDecimalActual As String
    Dim strMilesActual As String
    
    ' Obtener delimitadores actuales
    strDecimalActual = Application.DecimalSeparator
    strMilesActual = Application.ThousandsSeparator
    
    ' Verificar que coinciden con los deseados
    If strDecimalActual = vDelimitadorDecimal_HFM And strMilesActual = vDelimitadorMiles_HFM Then
        fun805_VerificarAplicacionDelimitadores = True
    Else
        fun805_VerificarAplicacionDelimitadores = False
    End If
    
    Exit Function

ErrorHandler_fun805:
    fun805_VerificarAplicacionDelimitadores = False
End Function

' =============================================================================
' FUNCI�N: fun806_RestaurarConfiguracion
' PROP�SITO: Restaura la configuraci�n anterior en caso de error
' FECHA: 2025-05-26 15:11:21 UTC
' =============================================================================
Private Sub fun806_RestaurarConfiguracion(ByVal strDecimalAnterior As String, ByVal strMilesAnterior As String)
    On Error Resume Next
    
    ' Restaurar delimitador decimal anterior
    Application.DecimalSeparator = strDecimalAnterior
    
    ' Restaurar delimitador de miles anterior
    Application.ThousandsSeparator = strMilesAnterior
    
    ' Restaurar uso de separadores del sistema
    Application.UseSystemSeparators = True
    
    On Error GoTo 0
End Sub

' =============================================================================
' FUNCI�N: fun807_MostrarErrorDetallado
' PROP�SITO: Muestra informaci�n detallada del error ocurrido
' FECHA: 2025-05-26 15:11:21 UTC
' =============================================================================
Private Sub fun807_MostrarErrorDetallado(ByVal strFuncion As String, ByVal strTipoError As String, _
                                        ByVal lngLinea As Long, ByVal lngNumeroError As Long, _
                                        ByVal strDescripcionError As String)
    
    Dim strMensajeError As String
    
    ' Construir mensaje de error detallado
    strMensajeError = "ERROR EN FUNCI�N DE DELIMITADORES" & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Funci�n: " & strFuncion & vbCrLf
    strMensajeError = strMensajeError & "Tipo de Error: " & strTipoError & vbCrLf
    strMensajeError = strMensajeError & "L�nea Aproximada: " & CStr(lngLinea) & vbCrLf
    strMensajeError = strMensajeError & "N�mero de Error VBA: " & CStr(lngNumeroError) & vbCrLf
    strMensajeError = strMensajeError & "Descripci�n: " & strDescripcionError & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Fecha/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Mostrar mensaje de error
    MsgBox strMensajeError, vbCritical, "Error en F004_Forzar_Delimitadores_en_Excel"
    
End Sub

