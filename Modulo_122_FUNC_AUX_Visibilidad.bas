Attribute VB_Name = "Modulo_122_FUNC_AUX_Visibilidad"
'===============================================================================
' DOCUMENTACI�N DE LA FUNCI�N
'===============================================================================
' Fecha y Hora de Creaci�n: 2025-05-26 17:04:54 UTC
' Autor: david-joaquin-corredera-de-colsa
' Funci�n: GestionarVisibilidadHoja
' Descripci�n: Funci�n que gestiona la existencia y visibilidad de una hoja de Excel
'
' RESUMEN EXHAUSTIVO DE PASOS:
' 1. Recibe como par�metro el nombre de una hoja (String)
' 2. Verifica si la hoja existe en el libro actual
' 3. Si la hoja NO existe:
'    - Crea una nueva hoja con el nombre especificado
'    - Contin�a al siguiente paso
' 4. Si la hoja S� existe:
'    - Comprueba el estado de visibilidad de la hoja
' 5. Si la hoja est� oculta (Hidden o VeryHidden):
'    - Hace la hoja visible
'    - Retorna informaci�n del cambio realizado
' 6. Si la hoja ya es visible:
'    - Termina la funci�n sin realizar cambios
'    - Retorna informaci�n del estado actual
'
' COMPATIBILIDAD:
' - VBA para Excel 365, Excel 97, Excel 2003
' - Compatible con libros en OneDrive, SharePoint, Teams
' - No utiliza clases ni caracter�sticas avanzadas de VBA
'
' PAR�METROS:
' - nombreHoja (String): Nombre de la hoja a gestionar
'
' VALOR DE RETORNO:
' - String: Mensaje descriptivo de la acci�n realizada
'
' CONTROL DE ERRORES:
' - Manejo exhaustivo de errores con informaci�n detallada
' - Identificaci�n de funci�n, tipo de error y l�nea
'===============================================================================

Public Function GestionarVisibilidadHoja(nombreHoja As String) As String
    ' Variables de control de errores
    Dim errorFuncion As String
    Dim errorTipo As String
    Dim errorLinea As String
    
    ' Variables de trabajo
    Dim ws As Worksheet
    Dim hojaExiste As Boolean
    Dim estadoVisibilidad As String
    Dim resultado As String
    
    ' Inicializaci�n de control de errores
    errorFuncion = "GestionarVisibilidadHoja"
    On Error GoTo ControlError
    
    ' PASO 1: Validaci�n del par�metro de entrada
    errorLinea = "Validaci�n par�metro nombreHoja"
    If Len(Trim(nombreHoja)) = 0 Then
        resultado = "ERROR: El nombre de la hoja no puede estar vac�o"
        GestionarVisibilidadHoja = resultado
        Exit Function
    End If
    
    ' PASO 2: Verificar si la hoja existe utilizando funci�n auxiliar
    errorLinea = "Verificaci�n existencia de hoja"
    hojaExiste = fun801_VerificarExistenciaHoja(nombreHoja)
    
    ' PASO 3: Procesar seg�n existencia de la hoja
    If Not hojaExiste Then
        ' PASO 3A: La hoja NO existe - Crear nueva hoja
        errorLinea = "Creaci�n de nueva hoja"
        resultado = fun802_CrearNuevaHoja(nombreHoja)
        
        ' Verificar que la creaci�n fue exitosa
        If InStr(resultado, "ERROR") > 0 Then
            GestionarVisibilidadHoja = resultado
            Exit Function
        End If
        
        ' Establecer referencia a la hoja reci�n creada
        errorLinea = "Establecer referencia a hoja creada"
        Set ws = ThisWorkbook.Worksheets(nombreHoja)
        
    Else
        ' PASO 3B: La hoja S� existe - Establecer referencia
        errorLinea = "Establecer referencia a hoja existente"
        Set ws = ThisWorkbook.Worksheets(nombreHoja)
    End If
    
    ' PASO 4: Verificar estado de visibilidad de la hoja
    errorLinea = "Verificaci�n estado visibilidad"
    estadoVisibilidad = fun803_ObtenerEstadoVisibilidad(ws)
    
    ' PASO 5: Procesar seg�n estado de visibilidad
    If estadoVisibilidad = "xlSheetHidden" Or estadoVisibilidad = "xlSheetVeryHidden" Then
        ' PASO 5A: La hoja est� oculta - Hacerla visible
        errorLinea = "Hacer hoja visible"
        resultado = fun804_HacerHojaVisible(ws, estadoVisibilidad)
        
    Else
        ' PASO 5B: La hoja ya es visible - Terminar funci�n
        errorLinea = "Hoja ya visible - Finalizaci�n"
        resultado = "INFO: La hoja '" & nombreHoja & "' ya es visible. No se requieren cambios."
    End If
    
    ' PASO 6: Retornar resultado final
    errorLinea = "Retorno de resultado"
    GestionarVisibilidadHoja = resultado
    
    ' Limpieza de referencias de objeto
    Set ws = Nothing
    Exit Function

ControlError:
    ' Control exhaustivo de errores
    errorTipo = "Error " & Err.Number & ": " & Err.Description
    resultado = "ERROR en funci�n " & errorFuncion & " - " & errorTipo & " - L�nea: " & errorLinea
    
    ' Limpieza de referencias en caso de error
    Set ws = Nothing
    
    ' Retornar informaci�n detallada del error
    GestionarVisibilidadHoja = resultado
End Function

'===============================================================================
' FUNCIONES AUXILIARES
'===============================================================================

'-------------------------------------------------------------------------------
' FUNCI�N AUXILIAR 801: Verificar existencia de hoja
'-------------------------------------------------------------------------------
Private Function fun801_VerificarExistenciaHoja(nombreHoja As String) As Boolean
    Dim ws As Worksheet
    Dim hojaEncontrada As Boolean
    
    On Error GoTo ErrorVerificacion
    
    ' Intentar obtener referencia a la hoja
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    hojaEncontrada = True
    
    ' Limpieza y retorno
    Set ws = Nothing
    fun801_VerificarExistenciaHoja = hojaEncontrada
    Exit Function
    
ErrorVerificacion:
    ' Si hay error, la hoja no existe
    hojaEncontrada = False
    Set ws = Nothing
    fun801_VerificarExistenciaHoja = hojaEncontrada
End Function

'-------------------------------------------------------------------------------
' FUNCI�N AUXILIAR 802: Crear nueva hoja
'-------------------------------------------------------------------------------
Private Function fun802_CrearNuevaHoja(nombreHoja As String) As String
    Dim nuevaHoja As Worksheet
    Dim resultado As String
    
    On Error GoTo ErrorCreacion
    
    ' Crear nueva hoja al final del libro
    Set nuevaHoja = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    ' Asignar nombre a la nueva hoja
    nuevaHoja.Name = nombreHoja
    
    ' Confirmar creaci�n exitosa
    resultado = "SUCCESS: Hoja '" & nombreHoja & "' creada exitosamente"
    
    ' Limpieza y retorno
    Set nuevaHoja = Nothing
    fun802_CrearNuevaHoja = resultado
    Exit Function
    
ErrorCreacion:
    ' Manejo de errores en creaci�n
    resultado = "ERROR al crear hoja '" & nombreHoja & "': " & Err.Description
    Set nuevaHoja = Nothing
    fun802_CrearNuevaHoja = resultado
End Function

'-------------------------------------------------------------------------------
' FUNCI�N AUXILIAR 803: Obtener estado de visibilidad
'-------------------------------------------------------------------------------
Private Function fun803_ObtenerEstadoVisibilidad(ws As Worksheet) As String
    Dim estadoVisibilidad As String
    
    On Error GoTo ErrorVisibilidad
    
    ' Determinar estado de visibilidad usando valores num�ricos para compatibilidad
    Select Case ws.Visible
        Case -1  ' xlSheetVisible
            estadoVisibilidad = "xlSheetVisible"
        Case 0   ' xlSheetHidden
            estadoVisibilidad = "xlSheetHidden"
        Case 2   ' xlSheetVeryHidden
            estadoVisibilidad = "xlSheetVeryHidden"
        Case Else
            estadoVisibilidad = "xlSheetVisible" ' Por defecto visible
    End Select
    
    fun803_ObtenerEstadoVisibilidad = estadoVisibilidad
    Exit Function
    
ErrorVisibilidad:
    ' En caso de error, asumir visible
    fun803_ObtenerEstadoVisibilidad = "xlSheetVisible"
End Function

'-------------------------------------------------------------------------------
' FUNCI�N AUXILIAR 804: Hacer hoja visible
'-------------------------------------------------------------------------------
Private Function fun804_HacerHojaVisible(ws As Worksheet, estadoAnterior As String) As String
    Dim resultado As String
    
    On Error GoTo ErrorHacerVisible
    
    ' Hacer la hoja visible usando valor num�rico para compatibilidad
    ws.Visible = -1  ' xlSheetVisible
    
    ' Confirmar cambio exitoso
    resultado = "SUCCESS: Hoja '" & ws.Name & "' cambiada de " & estadoAnterior & " a visible"
    
    fun804_HacerHojaVisible = resultado
    Exit Function
    
ErrorHacerVisible:
    ' Manejo de errores al hacer visible
    resultado = "ERROR al hacer visible la hoja '" & ws.Name & "': " & Err.Description
    fun804_HacerHojaVisible = resultado
End Function

'===============================================================================
' EJEMPLO DE USO DESDE UN SUB
'===============================================================================
Public Sub EjemploUsoGestionarHoja()
    '
    ' RESUMEN EXHAUSTIVO DE PASOS DE ESTE SUB DE EJEMPLO:
    ' 1. Declara variable para almacenar resultado
    ' 2. Llama a la funci�n GestionarHoja con nombre de hoja espec�fico
    ' 3. Muestra el resultado en un MsgBox para verificaci�n
    ' 4. Maneja cualquier error que pueda ocurrir durante la ejecuci�n
    '
    
    Dim resultadoOperacion As String
    
    On Error GoTo ErrorEjemplo
    
    ' Llamar a la funci�n con el nombre de hoja deseado
    resultadoOperacion = GestionarHoja("MiHojaDeDatos")
    
    ' Mostrar resultado de la operaci�n
    MsgBox "Resultado: " & resultadoOperacion, vbInformation, "Gesti�n de Hoja"
    
    Exit Sub
    
ErrorEjemplo:
    MsgBox "Error en EjemploUsoGestionarHoja: " & Err.Description, vbCritical, "Error"
End Sub
