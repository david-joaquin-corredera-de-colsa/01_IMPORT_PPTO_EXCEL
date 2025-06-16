Attribute VB_Name = "Modulo_800_Navegac_Y_Limpieza"

' =============================================================================
' MODULO: Modulo_Navegacion_Y_Limpieza.bas
' PROYECTO: IMPORTAR_DATOS_PRESUPUESTO
' AUTOR: david-joaquin-corredera-de-colsa
' FECHA CREACION: 2025-06-03 13:54:50 UTC
' FECHA ACTUALIZACION: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Modulo para navegacion inicial, limpieza de hojas historicas e inventario
' COMPATIBILIDAD: Excel 97, Excel 2003, Excel 2007, Excel 365
' REPOSITORIO: OneDrive, SharePoint, Teams compatible
' =============================================================================

Option Explicit

' Variables globales del modulo

Public Function F010_Abrir_Hoja_Inicial() As Boolean

    ' =============================================================================
    ' FUNCION: F010_Abrir_Hoja_Inicial
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Funcion para navegar a la hoja inicial del libro
    ' PARAMETROS: Ninguno
    ' RETORNO: Integer (0=exito, >0=error)
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Configurar la variable vHojaInicial con "00_Ejecutar_Procesos"
    ' 2. Verificar que el libro de trabajo este disponible
    ' 3. Buscar la hoja especificada en la coleccion de hojas del libro
    ' 4. Si la hoja existe, activarla y posicionarse en celda A1
    ' 5. Si la hoja no existe, retornar codigo de error
    ' 6. Retornar codigo de resultado

    On Error GoTo ErrorHandler
    
'    Dim vResultado As Integer
    Dim vHojaEncontrada As Boolean
    Dim vContadorHojas As Integer
    Dim vNombreHojaActual As String
    Dim vLineaError As Integer
    
'    vResultado = 0
    vHojaEncontrada = False
    vContadorHojas = 0
    vLineaError = 10
    
    ' Paso 1: Configurar la variable vHojaInicial con "00_Ejecutar_Procesos"
    Dim vHojaInicial As String
    vHojaInicial = CONST_HOJA_EJECUTAR_PROCESOS
    
    
    vLineaError = 20
    
    ' Paso 2: Verificar que el libro de trabajo este disponible
    'vLineaError = 30
    'If ThisWorkbook Is Nothing Then
    '    F010_Abrir_Hoja_Inicial = False ' Error: Libro de trabajo no disponible
    '    GoTo ErrorHandler
    'End If
    
    ' Paso 3: Buscar la hoja especificada en la coleccion de hojas del libro
    vLineaError = 40
    For vContadorHojas = 1 To ThisWorkbook.Worksheets.Count
        vNombreHojaActual = ThisWorkbook.Worksheets(vContadorHojas).Name
        If StrComp(vNombreHojaActual, vHojaInicial, vbTextCompare) = 0 Then
            vHojaEncontrada = True
            Exit For
        End If
    Next vContadorHojas
    
    ' Paso 4: Si la hoja existe, activarla y posicionarse en celda A1
    vLineaError = 50
    If vHojaEncontrada Then
        ThisWorkbook.Worksheets(vHojaInicial).Activate
        vLineaError = 55
        ThisWorkbook.Worksheets(vHojaInicial).Range("A1").Select
        F010_Abrir_Hoja_Inicial = True ' Exito
    Else
        ' Paso 5: Si la hoja no existe, retornar codigo de error
        F010_Abrir_Hoja_Inicial = False ' Error: Hoja no encontrada
    End If
    
    ' Paso 6: Retornar codigo de resultado
    'F010_Abrir_Hoja_Inicial = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F010_Abrir_Hoja_Inicial" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Hoja objetivo: " & vHojaInicial
    
    MsgBox vMensajeError, vbCritical, "Error F010_Abrir_Hoja_Inicial"
        
    F010_Abrir_Hoja_Inicial = False 'OJO
    
End Function



Public Function F011_Limpieza_Hojas_Historicas() As Boolean

    ' =============================================================================
    ' FUNCION: F011_Limpieza_Hojas_Historicas
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Funcion para limpiar hojas historicas segun criterios especificos
    ' PARAMETROS: Ninguno
    ' RETORNO: Integer (0=exito, >0=error)
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables de control
    ' 1.5. Configurar visibilidad de hojas especificas
    ' 2. Primera pasada - recopilar hojas Import_Envio_
    ' 3. Ordenar hojas Import_Envio_ lexicograficamente
    ' 4. Segunda pasada - aplicar reglas de limpieza especificas
    ' 5. Gestionar hojas Import_Envio_ con logica de ordenamiento lexicografico
    ' 6. Retornar codigo de resultado
    
    
    On Error GoTo ErrorHandler
    
    'Dim vResultado As Integer
    Dim vContadorHojas As Integer
    Dim vNombreHoja As String
    Dim vLineaError As Integer
    Dim vTotalHojas As Integer
    Dim vHojasEnvio() As String
    Dim vContadorEnvio As Integer
    Dim vNumHojasEnvio As Integer
    Dim i As Integer, j As Integer
    Dim vTempNombre As String
    
    Dim strFuncion As String
    strFuncion = "F011_Limpieza_Hojas_Historicas" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F011_Limpieza_Hojas_Historicas"
    'MsgBox "strFuncion = " & strFuncion
    
    'vResultado = 0
    vLineaError = 10
    vContadorEnvio = 0
    vNumHojasEnvio = 0
    
    ' Paso 1: Inicializar variables de control
    vLineaError = 20
    vTotalHojas = ThisWorkbook.Worksheets.Count
    
    ' Paso 1.5: Configurar visibilidad de Hojas Tecnicas
    vLineaError = 25
    ThisWorkbook.Worksheets(CONST_HOJA_EJECUTAR_PROCESOS).Visible = CONST_HOJA_EJECUTAR_PROCESOS_VISIBLE
    ThisWorkbook.Worksheets(CONST_HOJA_INVENTARIO).Visible = CONST_HOJA_INVENTARIO_VISIBLE
    ThisWorkbook.Worksheets(CONST_HOJA_LOG).Visible = CONST_HOJA_LOG_VISIBLE
    ThisWorkbook.Worksheets(CONST_HOJA_USERNAME).Visible = CONST_HOJA_USERNAME_VISIBLE
    ThisWorkbook.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES).Visible = CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE
    ThisWorkbook.Worksheets(CONST_HOJA_REPORT_PL).Visible = CONST_HOJA_REPORT_PL_VISIBLE
    ThisWorkbook.Worksheets(CONST_HOJA_REPORT_PL_AH).Visible = CONST_HOJA_REPORT_PL_AH_VISIBLE
        
    ' Redimensionar array para hojas Import_Envio_ (estimacion maxima)
    ReDim vHojasEnvio(1 To vTotalHojas)
    
    ' Paso 2: Primera pasada - recopilar hojas Import_Envio_ con prefijo "Import_Envio_"
    vLineaError = 30
    For vContadorHojas = vTotalHojas To 1 Step -1
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        If Left(UCase(vNombreHoja), Len(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO)) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            vNumHojasEnvio = vNumHojasEnvio + 1
            vHojasEnvio(vNumHojasEnvio) = vNombreHoja
        End If
    Next vContadorHojas
    
    ' Paso 3: Ordenar hojas Import_Envio_ lexicograficamente (bubble sort compatible Excel 97)
    vLineaError = 40
    If vNumHojasEnvio > 1 Then
        For i = 1 To vNumHojasEnvio - 1
            For j = 1 To vNumHojasEnvio - i
                If StrComp(vHojasEnvio(j), vHojasEnvio(j + 1), vbTextCompare) < 0 Then
                    vTempNombre = vHojasEnvio(j)
                    vHojasEnvio(j) = vHojasEnvio(j + 1)
                    vHojasEnvio(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Paso 4: Segunda pasada - aplicar reglas de limpieza especificas
    vLineaError = 50
    For vContadorHojas = vTotalHojas To 1 Step -1
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        ' Para cada hoja aplicar reglas de limpieza especificas
        vLineaError = 60
        
        ' Regla: Hojas protegidas - no hacer nada
        If fun805_Es_Hoja_Protegida(vNombreHoja) Then
            ' No hacer nada con estas hojas
            
        ' Regla: Eliminar hojas Import_Working_ con prefijo "Import_Working_"
        ElseIf Left(UCase(vNombreHoja), Len(CONST_PREFIJO_HOJA_IMPORTACION_WORKING)) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_WORKING) Then
            vLineaError = 70
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Import_Comprob_ con prefijo "Import_Comprob_"
        ElseIf Left(UCase(vNombreHoja), Len(CONST_PREFIJO_HOJA_IMPORTACION_COMPROBACION)) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_COMPROBACION) Then
            vLineaError = 80
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Import_ con prefijo solo "Import_"
        '   Ya no va a considerar las que tienen prefijo CONST_PREFIJO_HOJA_IMPORTACION_WORKING ("Import_Working_") porque estaba en ramas anteriores del If
        '   Ya no va a considerar las que tienen prefijo CONST_PREFIJO_HOJA_IMPORTACION_COMPROBACION ("Import_Comprob_") porque estaba en ramas anteriores del If
        '   Ya solo tengo que comprobar que ademas de empezar por el prefijo CONST_PREFIJO_HOJA_IMPORTACION ("Import_"),
        '       no comience por el prefijo CONST_PREFIJO_HOJA_IMPORTACION_ENVIO ("Import_Envio_")
        ElseIf ((Left(vNombreHoja, Len(CONST_PREFIJO_HOJA_IMPORTACION)) = CONST_PREFIJO_HOJA_IMPORTACION) And _
                (Not (Left(vNombreHoja, Len(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO)) = CONST_PREFIJO_HOJA_IMPORTACION_ENVIO))) Then
            vLineaError = 90
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas con prefijo CONST_PREFIJO_HOJA_X_BORRAR_ENVIO_PREVIO (normalmente "Del_Prev_Envio_")h
        ElseIf Left(vNombreHoja, 15) = CONST_PREFIJO_HOJA_X_BORRAR_ENVIO_PREVIO Then
            vLineaError = 100
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Gestionar hojas Import_Envio_ con prefijo "Import_Envio_"
        ElseIf Left(UCase(vNombreHoja), Len(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO)) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            vLineaError = 110
            Call fun807_Gestionar_Hoja_Envio(vNombreHoja, vHojasEnvio, vNumHojasEnvio)
        End If
    Next vContadorHojas
    
    ' Paso 6: Retornar codigo de resultado
    F011_Limpieza_Hojas_Historicas = True
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F011_Limpieza_Hojas_Historicas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Hoja procesando: " & vNombreHoja
    
    MsgBox vMensajeError, vbCritical, "Error F011_Limpieza_Hojas_Historicas"
        
    F011_Limpieza_Hojas_Historicas = False
    
End Function
Public Function Function_Return_Integer_to_Boolean(vInteger As Integer) As Boolean
    '******************************************************************************
    ' Módulo: Function_Return_Integer_to_Boolean
    ' Fecha y Hora de Creación: 2025-06-09 09:10:01 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Convierte un valor entero a un valor booleano siguiendo una lógica específica
    ' donde 0 se considera verdadero (True) y cualquier otro valor se considera
    ' falso (False). Esta función implementa una lógica inversa a la conversión
    ' booleana estándar de VBA.
    '
    ' Parámetros:
    ' - vInteger (Integer): Valor entero a convertir a booleano
    '
    ' Valor de Retorno:
    ' - Boolean: True si el valor de entrada es 0, False para cualquier otro valor
    '
    ' Lógica de Conversión:
    ' - Input: 0 ? Output: True
    ' - Input: cualquier otro número ? Output: False
    '
    ' Casos de Uso Típicos:
    ' - Validación de códigos de error (donde 0 indica éxito)
    ' - Conversión de flags numéricos a booleanos
    ' - Procesamiento de datos donde 0 representa un estado "activo" o "válido"
    '
    ' Ejemplos de Uso:
    ' Dim resultado As Boolean
    ' resultado = Function_Return_Integer_to_Boolean(0)    ' Devuelve True
    ' resultado = Function_Return_Integer_to_Boolean(1)    ' Devuelve False
    ' resultado = Function_Return_Integer_to_Boolean(-5)   ' Devuelve False
    '
    ' Notas Importantes:
    ' ?? Esta función implementa una lógica inversa a la conversión booleana
    ' estándar de VBA, donde normalmente 0 equivale a False y cualquier valor
    ' diferente de 0 equivale a True.
    '
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' Versión: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Inicialización
    strFuncion = "Function_Return_Integer_to_Boolean" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Function_Return_Integer_to_Boolean"
    Function_Return_Integer_to_Boolean = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' Lógica principal de conversión
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    If vInteger = 0 Then
        Function_Return_Integer_to_Boolean = True
    Else
        Function_Return_Integer_to_Boolean = False
    End If
    
    Exit Function

GestorErrores:
    ' Manejo de errores con información detallada
    Dim strMensajeError As String
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Valor de entrada: " & vInteger
    
    ' Log del error para debugging
    Debug.Print strMensajeError
    
    ' Retornar False en caso de error
    Function_Return_Integer_to_Boolean = False
End Function



Public Function F012_Inventariar_Hojas() As Boolean

    ' =============================================================================
    ' FUNCION: F012_Inventariar_Hojas
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Funcion para crear inventario completo de todas las hojas del libro
    ' PARAMETROS: Ninguno
    ' RETORNO: Integer (0=exito, >0=error)
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Verificar existencia de hoja "01_Inventario"
    ' 2. Borrar contenido y formatos de la hoja "01_Inventario"
    ' 3. Crear encabezados en linea 2 con formato especifico
    ' 4. Recorrer todas las hojas y recopilar informacion completa
    ' 5. Crear enlaces (hyperlinks) para cada hoja
    ' 6. Aplicar formato segun visibilidad de cada hoja
    ' 7. Buscar fichero fuente en hoja "02_Log"
    ' 8. Ordenar el listado alfabeticamente
    ' 9. Asegurar visibilidad de hoja "01_Inventario"

    On Error GoTo ErrorHandler
    
    'Dim vResultado As Integer
    Dim vLineaError As Integer
    Dim vHojaInventario As Worksheet
    Dim vContadorHojas As Integer
    Dim vFilaActual As Integer
    Dim vNombreHoja As String
    Dim vRangoOrdenar As Range
    Dim vEsVisible As Boolean
    Dim vFicheroFuente As String
    
    Dim strFuncion As String
    strFuncion = "F012_Inventariar_Hojas" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F012_Inventariar_Hojas"
    'MsgBox "strFuncion = " & strFuncion

    
    'vResultado = 0
    vLineaError = 10
    vFilaActual = 3
    
    ' Paso 1: Verificar existencia de hoja "01_Inventario"
    vLineaError = 20
    Set vHojaInventario = ThisWorkbook.Worksheets(CONST_HOJA_INVENTARIO)
    If vHojaInventario Is Nothing Then
        F012_Inventariar_Hojas = False
        GoTo ErrorHandler
    End If
    
    ' Paso 2: Borrar contenido y formatos de la hoja "01_Inventario"
    vLineaError = 30
    vHojaInventario.Cells.Clear
    vHojaInventario.Cells.ClearFormats
    
    ' Paso 3: Crear encabezados en linea CONST_INVENTARIO_FILA_HEADERS = 2
    vLineaError = 40
    vHojaInventario.Cells(CONST_INVENTARIO_FILA_HEADERS, CONST_INVENTARIO_COLUMNA_NOMBRE).Value = CONST_INVENTARIO_HEADER_NOMBRE    'Columna CONST_INVENTARIO_COLUMNA_NOMBRE = 2 y CONST_INVENTARIO_HEADER_NOMBRE = "Nombre de la Hoja"
    vHojaInventario.Cells(CONST_INVENTARIO_FILA_HEADERS, CONST_INVENTARIO_COLUMNA_LINK).Value = CONST_INVENTARIO_HEADER_LINK        'Columna CONST_INVENTARIO_COLUMNA_LINK = 3 y CONST_INVENTARIO_HEADER_LINK = "Link a la Hoja"
    vHojaInventario.Cells(CONST_INVENTARIO_FILA_HEADERS, CONST_INVENTARIO_COLUMNA_VISIBLE).Value = CONST_INVENTARIO_HEADER_VISIBLE  'Columna CONST_INVENTARIO_COLUMNA_VISIBLE = 4 y CONST_INVENTARIO_HEADER_VISIBLE = "Visible/Oculta"
    vHojaInventario.Cells(CONST_INVENTARIO_FILA_HEADERS, CONST_INVENTARIO_COLUMNA_FICHERO).Value = CONST_INVENTARIO_HEADER_FICHERO  'Columna CONST_INVENTARIO_COLUMNA_FICHERO = 5 y CONST_INVENTARIO_HEADER_FICHERO = "Fichero Fuente"
    
    ' Paso 3.1: Aplicar formato a encabezados
    vLineaError = 45
    Call fun803_Aplicar_Formato_Inventario_Encabezados(vHojaInventario)
    
    ' Paso 4: Recorrer todas las hojas y recopilar informacion completa
    vLineaError = 50
    For vContadorHojas = 1 To ThisWorkbook.Worksheets.Count
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        ' Paso 4.1: Escribir nombre de la hoja
        vLineaError = 60
        vHojaInventario.Cells(vFilaActual, CONST_INVENTARIO_COLUMNA_NOMBRE).Value = vNombreHoja 'Columna CONST_INVENTARIO_COLUMNA_NOMBRE = 2
        
        ' Paso 4.2: Crear enlaces (hyperlinks) para cada hoja
        vLineaError = 70
        Call fun809_Crear_Enlace_Hoja(vHojaInventario, vFilaActual, CONST_INVENTARIO_COLUMNA_LINK, vNombreHoja) 'Columna CONST_INVENTARIO_COLUMNA_LINK = 3
        
        ' Paso 4.3: Determinar si la hoja es visible
        vLineaError = 80
        vEsVisible = (ThisWorkbook.Worksheets(vNombreHoja).Visible = xlSheetVisible)
        
        ' Paso 4.4: Aplicar formato segun visibilidad
        vLineaError = 90
        Call fun804_Aplicar_Formato_Inventario_Fila(vHojaInventario, vFilaActual, vEsVisible)
        
        ' Paso 4.5: Buscar fichero fuente en hoja "02_Log"
        vLineaError = 100
        vFicheroFuente = fun802_Buscar_Fichero_Fuente_En_Log(vNombreHoja)
        vHojaInventario.Cells(vFilaActual, CONST_INVENTARIO_COLUMNA_FICHERO).Value = vFicheroFuente 'Columna CONST_INVENTARIO_COLUMNA_FICHERO = 5
        
        vFilaActual = vFilaActual + 1
    Next vContadorHojas
    
    ' Paso 5: Ordenar el listado alfabeticamente (compatible Excel 97)
    vLineaError = 110
    If vFilaActual > 3 Then
        Set vRangoOrdenar = vHojaInventario.Range("B3:E" & (vFilaActual - 1))
        'vRangoOrdenar.Sort Key1:=vHojaInventario.Range("B3"), Order1:=xlAscending, Header:=xlNo
        vRangoOrdenar.Sort Key1:=vHojaInventario.Range("D3"), Order1:=xlAscending, Key2:=vHojaInventario.Range("B3"), Order2:=xlAscending, Header:=xlNo
    End If
    
    ' Paso 6: Asegurar visibilidad de hoja "01_Inventario"
    vLineaError = 120
    ThisWorkbook.Worksheets(CONST_HOJA_INVENTARIO).Visible = CONST_HOJA_INVENTARIO_VISIBLE
    
    ' Ajustar columnas automaticamente
    vLineaError = 125
    vHojaInventario.Columns("B:E").AutoFit
    
    F012_Inventariar_Hojas = True
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F012_Inventariar_Hojas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description
    
    MsgBox vMensajeError, vbCritical, "Error F012_Inventariar_Hojas"
    
    F012_Inventariar_Hojas = False
    
End Function


' =============================================================================
' FUNCION AUXILIAR: fun802_Buscar_Fichero_Fuente_En_Log
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Busca el fichero fuente de una hoja en el log
' PARAMETROS: vNombreHoja (String)
' RETORNO: String (nombre del fichero fuente o "")
' =============================================================================
Public Function fun802_Buscar_Fichero_Fuente_En_Log(vNombreHoja As String) As String
    
    On Error GoTo ErrorHandler
    
    Dim vHojaLog As Worksheet
    Dim vUltimaFila As Long
    Dim i As Long
    Dim vValorColumnaD As String
    Dim vValorColumnaE As String
    Dim vBuscarTexto As String
    
    fun802_Buscar_Fichero_Fuente_En_Log = ""
    
    ' Obtener referencia a la hoja "02_Log"
    Set vHojaLog = Nothing
    ' Cogemos la hoja cuyo nombre esta almacenado en la constante CONST_HOJA_LOG y su valor es "02_Log"
    Set vHojaLog = ThisWorkbook.Worksheets(CONST_HOJA_LOG)
    If vHojaLog Is Nothing Then Exit Function
    
    ' Determinar texto a buscar basado en el nombre de la hoja prefijo "Import_Envio_"
    ' Donde CONST_PREFIJO_HOJA_IMPORTACION = "Import_" y CONST_PREFIJO_HOJA_IMPORTACION_ENVIO = "Import_Envio_"
    If Left(UCase(vNombreHoja), Len(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO)) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
        vBuscarTexto = CONST_PREFIJO_HOJA_IMPORTACION & Right(vNombreHoja, 15)
    Else
        vBuscarTexto = vNombreHoja
    End If
    
    ' Obtener ultima fila con datos
    vUltimaFila = vHojaLog.Cells(vHojaLog.Rows.Count, "D").End(xlUp).Row
    
    ' Recorrer las filas buscando la coincidencia
    For i = 1 To vUltimaFila
        vValorColumnaD = CStr(vHojaLog.Cells(i, 4).Value)
        vValorColumnaE = CStr(vHojaLog.Cells(i, 5).Value)
        
        ' Verificar condiciones: columna D diferente de "NA" y contiene "\"
        ' y columna E igual al texto buscado
        If StrComp(vValorColumnaD, "NA", vbTextCompare) <> 0 And _
           InStr(vValorColumnaD, "\") > 0 And _
           StrComp(vValorColumnaE, vBuscarTexto, vbTextCompare) = 0 Then
            fun802_Buscar_Fichero_Fuente_En_Log = vValorColumnaD
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    fun802_Buscar_Fichero_Fuente_En_Log = ""
    
End Function

Public Sub fun803_Aplicar_Formato_Inventario_Encabezados(vHojaInventario As Worksheet)
    
    ' =============================================================================
    ' FUNCION AUXILIAR: fun803_Aplicar_Formato_Inventario_Encabezados
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Aplica formato a los encabezados del inventario
    ' PARAMETROS: vHojaInventario (Worksheet)
    ' =============================================================================
    On Error GoTo ErrorHandler
    
    Dim vRangoEncabezados As Range
    
    ' Definir rango de encabezados (fila 2, columnas 2 a 5)
    Set vRangoEncabezados = vHojaInventario.Range("B2:E2")
    
    ' Aplicar formato de fondo negro
    vRangoEncabezados.Interior.Color = RGB(0, 0, 0)
    
    ' Aplicar formato de fuente blanca y negrita
    With vRangoEncabezados.Font
        .Color = RGB(255, 255, 255)
        .Bold = True
    End With
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

Public Sub fun804_Aplicar_Formato_Inventario_Fila(vHojaInventario As Worksheet, vFila As Integer, vEsVisible As Boolean)

    ' =============================================================================
    ' FUNCION AUXILIAR: fun804_Aplicar_Formato_Inventario_Fila
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Aplica formato a una fila del inventario segun visibilidad
    ' PARAMETROS: vHojaInventario (Worksheet), vFila (Integer), vEsVisible (Boolean)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vRangoFila As Range
    
    ' Definir rango de la fila (columnas 2 a 4)
    Set vRangoFila = vHojaInventario.Range("B" & vFila & ":D" & vFila)
    
    If vEsVisible Then
        ' Fila visible: sin color de fondo
        vRangoFila.Interior.ColorIndex = xlNone
        vHojaInventario.Cells(vFila, CONST_INVENTARIO_COLUMNA_VISIBLE).Value = CONST_INVENTARIO_TAG_VISIBLE  'Columna CONST_INVENTARIO_COLUMNA_VISIBLE = 4 y CONST_INVENTARIO_TAG_VISIBLE = "Visible/Oculta"
    Else
        ' Fila oculta: fondo gris medio
        vRangoFila.Interior.Color = RGB(128, 128, 128)
        vHojaInventario.Cells(vFila, CONST_INVENTARIO_COLUMNA_VISIBLE).Value = CONST_INVENTARIO_TAG_OCULTA         'Columna CONST_INVENTARIO_COLUMNA_VISIBLE = 4 y CONST_INVENTARIO_TAG_OCULTA = "OCULTA"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

Public Function fun805_Es_Hoja_Protegida(vNombreHoja As String) As Boolean
    
    ' =============================================================================
    ' FUNCION AUXILIAR: fun805_Es_Hoja_Protegida
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Verifica si una hoja esta en la lista de hojas protegidas
    ' PARAMETROS: vNombreHoja (String)
    ' RETORNO: Boolean (True=protegida, False=no protegida)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vHojasProtegidas(1 To 6) As String
    Dim i As Integer
    
    ' Lista de hojas protegidas
    vHojasProtegidas(1) = CONST_HOJA_EJECUTAR_PROCESOS ' "00_Ejecutar_Procesos"
    vHojasProtegidas(2) = CONST_HOJA_INVENTARIO ' "01_Inventario"
    vHojasProtegidas(3) = CONST_HOJA_USERNAME ' "05_Username"
    vHojasProtegidas(4) = CONST_HOJA_DELIMITADORES_ORIGINALES ' "06_Delimitadores_Originales"
    vHojasProtegidas(5) = CONST_HOJA_REPORT_PL ' "09_Report_PL"
    vHojasProtegidas(6) = CONST_HOJA_REPORT_PL_AH ' "10_Report_PL_AH"
    
    fun805_Es_Hoja_Protegida = False
    
    For i = 1 To 6
        If StrComp(vNombreHoja, vHojasProtegidas(i), vbTextCompare) = 0 Then
            fun805_Es_Hoja_Protegida = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    fun805_Es_Hoja_Protegida = False
    
End Function

Public Sub fun806_Eliminar_Hoja_Segura(vNombreHoja As String)
    
    ' =============================================================================
    ' SUB AUXILIAR: fun806_Eliminar_Hoja_Segura
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Elimina una hoja de forma segura con control de errores
    ' PARAMETROS: vNombreHoja (String)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vAlertas As Boolean
    
    ' Desactivar alertas para evitar confirmaciones
    vAlertas = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    ' Eliminar la hoja
    ThisWorkbook.Worksheets(vNombreHoja).Delete
    
    ' Restaurar alertas
    Application.DisplayAlerts = vAlertas
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = vAlertas
    ' No mostrar error, simplemente continuar
    
End Sub

Public Sub fun807_Gestionar_Hoja_Envio(vNombreHoja As String, vHojasEnvio() As String, vNumHojasEnvio As Integer)
    
    ' =============================================================================
    ' SUB AUXILIAR: fun807_Gestionar_Hoja_Envio
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Gestiona visibilidad de hojas Import_Envio_ segun antiguedad
    ' PARAMETROS: vNombreHoja (String), vHojasEnvio (Array), vNumHojasEnvio (Integer)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim vPosicion As Integer
    Dim vHoja As Worksheet
    
    ' Buscar posicion de la hoja en el array ordenado
    vPosicion = 0
    For i = 1 To vNumHojasEnvio
        If StrComp(vHojasEnvio(i), vNombreHoja, vbTextCompare) = 0 Then
            vPosicion = i
            Exit For
        End If
    Next i
    
    Set vHoja = ThisWorkbook.Worksheets(vNombreHoja)
    
    ' Si hay mas hojas que el limite y esta fuera del rango visible
    If vNumHojasEnvio > CONST_NUM_HOJAS_HCAS_VISIBLES_ENVIO And vPosicion > CONST_NUM_HOJAS_HCAS_VISIBLES_ENVIO Then
        vHoja.Visible = xlSheetHidden
    Else
        vHoja.Visible = xlSheetVisible
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub



Public Sub fun809_Crear_Enlace_Hoja(vHojaDestino As Worksheet, vFila As Integer, vColumna As Integer, vNombreHoja As String)

    ' =============================================================================
    ' SUB AUXILIAR: fun809_Crear_Enlace_Hoja
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Crea un hyperlink a una hoja especifica (compatible Excel 97)
    ' PARAMETROS: vHojaDestino, vFila, vColumna, vNombreHoja
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vCelda As Range
    Dim vDireccion As String
    
    Set vCelda = vHojaDestino.Cells(vFila, vColumna)
    vDireccion = "'" & vNombreHoja & "'!A1"
    
    ' Metodo compatible con Excel 97
    vCelda.Value = "Ir a " & vNombreHoja
    vCelda.Font.ColorIndex = 5 ' Azul
    vCelda.Font.Underline = xlUnderlineStyleSingle
    
    ' Crear hyperlink (Excel 97+ compatible)
    vHojaDestino.Hyperlinks.Add Anchor:=vCelda, Address:="", SubAddress:=vDireccion, TextToDisplay:="Ir a " & vNombreHoja
    
    Exit Sub
    
ErrorHandler:
    ' Si falla el hyperlink, al menos mostrar el texto
    vCelda.Value = vNombreHoja
    
End Sub





