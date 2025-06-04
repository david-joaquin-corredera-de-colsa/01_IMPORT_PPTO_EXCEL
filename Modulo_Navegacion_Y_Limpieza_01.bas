Attribute VB_Name = "Modulo_Navegacion_Y_Limpieza_01"

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
Public vHojaInicial As String
Public Const vNum_Hojas_Hcas_Visibles_Envio As Integer = 5

' =============================================================================
' FUNCION: F010_Abrir_Hoja_Inicial
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Funcion para navegar a la hoja inicial del libro
' PARAMETROS: Ninguno
' RETORNO: Integer (0=exito, >0=error)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function F010_Abrir_Hoja_Inicial() As Integer

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Configurar la variable vHojaInicial con "00_Ejecutar_Procesos"
    ' 2. Verificar que el libro de trabajo este disponible
    ' 3. Buscar la hoja especificada en la coleccion de hojas del libro
    ' 4. Si la hoja existe, activarla y posicionarse en celda A1
    ' 5. Si la hoja no existe, retornar codigo de error
    ' 6. Retornar codigo de resultado

    On Error GoTo ErrorHandler
    
    Dim vResultado As Integer
    Dim vHojaEncontrada As Boolean
    Dim vContadorHojas As Integer
    Dim vNombreHojaActual As String
    Dim vLineaError As Integer
    
    vResultado = 0
    vHojaEncontrada = False
    vContadorHojas = 0
    vLineaError = 10
    
    ' Paso 1: Configurar la variable vHojaInicial con "00_Ejecutar_Procesos"
    vHojaInicial = "00_Ejecutar_Procesos"
    vLineaError = 20
    
    ' Paso 2: Verificar que el libro de trabajo este disponible
    vLineaError = 30
    If ThisWorkbook Is Nothing Then
        vResultado = 1001 ' Error: Libro de trabajo no disponible
        GoTo ErrorHandler
    End If
    
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
        vResultado = 0 ' Exito
    Else
        ' Paso 5: Si la hoja no existe, retornar codigo de error
        vResultado = 1002 ' Error: Hoja no encontrada
    End If
    
    ' Paso 6: Retornar codigo de resultado
    F010_Abrir_Hoja_Inicial = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F010_Abrir_Hoja_Inicial" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Hoja objetivo: " & vHojaInicial
    
    MsgBox vMensajeError, vbCritical, "Error F010_Abrir_Hoja_Inicial"
    
    If vResultado = 0 Then
        vResultado = 9999 ' Error no especificado
    End If
    
    F010_Abrir_Hoja_Inicial = vResultado
    
End Function

' =============================================================================
' FUNCION: F011_Limpieza_Hojas_Historicas
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Funcion para limpiar hojas historicas segun criterios especificos
' PARAMETROS: Ninguno
' RETORNO: Integer (0=exito, >0=error)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function F011_Limpieza_Hojas_Historicas() As Integer

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables de control
    ' 1.5. Configurar visibilidad de hojas especificas
    ' 2. Primera pasada - recopilar hojas Import_Envio_
    ' 3. Ordenar hojas Import_Envio_ lexicograficamente
    ' 4. Segunda pasada - aplicar reglas de limpieza especificas
    ' 5. Gestionar hojas Import_Envio_ con logica de ordenamiento lexicografico
    ' 6. Retornar codigo de resultado

    On Error GoTo ErrorHandler
    
    Dim vResultado As Integer
    Dim vContadorHojas As Integer
    Dim vNombreHoja As String
    Dim vLineaError As Integer
    Dim vTotalHojas As Integer
    Dim vHojasEnvio() As String
    Dim vContadorEnvio As Integer
    Dim vNumHojasEnvio As Integer
    Dim i As Integer, j As Integer
    Dim vTempNombre As String
    
    vResultado = 0
    vLineaError = 10
    vContadorEnvio = 0
    vNumHojasEnvio = 0
    
    ' Paso 1: Inicializar variables de control
    vLineaError = 20
    vTotalHojas = ThisWorkbook.Worksheets.Count
    
    ' Paso 1.5: Configurar visibilidad de hojas especificas
    vLineaError = 25
    For vContadorHojas = 1 To vTotalHojas
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        Call fun801_Configurar_Visibilidad_Hojas_Especificas(vNombreHoja)
    Next vContadorHojas
    
    ' Redimensionar array para hojas Import_Envio_ (estimacion maxima)
    ReDim vHojasEnvio(1 To vTotalHojas)
    
    ' Paso 2: Primera pasada - recopilar hojas Import_Envio_
    vLineaError = 30
    For vContadorHojas = vTotalHojas To 1 Step -1
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        If Left(vNombreHoja, 13) = "Import_Envio_" Then
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
            
        ' Regla: Eliminar hojas Import_Working_
        ElseIf Left(vNombreHoja, 15) = "Import_Working_" Then
            vLineaError = 70
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Import_Comprob_
        ElseIf Left(vNombreHoja, 15) = "Import_Comprob_" Then
            vLineaError = 80
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Import_ con longitud 22
        ElseIf Left(vNombreHoja, 7) = "Import_" And Len(vNombreHoja) = 22 Then
            vLineaError = 90
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Del_Prev_Envio_
        ElseIf Left(vNombreHoja, 15) = "Del_Prev_Envio_" Then
            vLineaError = 100
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Gestionar hojas Import_Envio_
        ElseIf Left(vNombreHoja, 13) = "Import_Envio_" Then
            vLineaError = 110
            Call fun807_Gestionar_Hoja_Envio(vNombreHoja, vHojasEnvio, vNumHojasEnvio)
        End If
    Next vContadorHojas
    
    ' Paso 6: Retornar codigo de resultado
    F011_Limpieza_Hojas_Historicas = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F011_Limpieza_Hojas_Historicas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Hoja procesando: " & vNombreHoja
    
    MsgBox vMensajeError, vbCritical, "Error F011_Limpieza_Hojas_Historicas"
    
    If vResultado = 0 Then
        vResultado = 2999 ' Error no especificado en limpieza
    End If
    
    F011_Limpieza_Hojas_Historicas = vResultado
    
End Function
Public Function Function_Return_Integer_to_Boolean(vInteger As Integer) As Boolean
    If vInteger = 0 Then
        Function_Return_Integer_to_Boolean = True
    Else
        Function_Return_Integer_to_Boolean = False
    End If
End Function

' =============================================================================
' FUNCION: F012_Inventariar_Hojas
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Funcion para crear inventario completo de todas las hojas del libro
' PARAMETROS: Ninguno
' RETORNO: Integer (0=exito, >0=error)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function F012_Inventariar_Hojas() As Integer

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
    
    Dim vResultado As Integer
    Dim vLineaError As Integer
    Dim vHojaInventario As Worksheet
    Dim vContadorHojas As Integer
    Dim vFilaActual As Integer
    Dim vNombreHoja As String
    Dim vRangoOrdenar As Range
    Dim vEsVisible As Boolean
    Dim vFicheroFuente As String
    
    vResultado = 0
    vLineaError = 10
    vFilaActual = 3
    
    ' Paso 1: Verificar existencia de hoja "01_Inventario"
    vLineaError = 20
    Set vHojaInventario = fun808_Obtener_Hoja_Inventario()
    If vHojaInventario Is Nothing Then
        vResultado = 3001 ' Error: No se pudo acceder a hoja inventario
        GoTo ErrorHandler
    End If
    
    ' Paso 2: Borrar contenido y formatos de la hoja "01_Inventario"
    vLineaError = 30
    vHojaInventario.Cells.Clear
    vHojaInventario.Cells.ClearFormats
    
    ' Paso 3: Crear encabezados en linea 2
    vLineaError = 40
    vHojaInventario.Cells(2, 2).Value = "Nombre de la Hoja"
    vHojaInventario.Cells(2, 3).Value = "Link a la Hoja"
    vHojaInventario.Cells(2, 4).Value = "Visible/Oculta"
    vHojaInventario.Cells(2, 5).Value = "Fichero Fuente"
    
    ' Paso 3.1: Aplicar formato a encabezados
    vLineaError = 45
    Call fun803_Aplicar_Formato_Inventario_Encabezados(vHojaInventario)
    
    ' Paso 4: Recorrer todas las hojas y recopilar informacion completa
    vLineaError = 50
    For vContadorHojas = 1 To ThisWorkbook.Worksheets.Count
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        ' Paso 4.1: Escribir nombre de la hoja
        vLineaError = 60
        vHojaInventario.Cells(vFilaActual, 2).Value = vNombreHoja
        
        ' Paso 4.2: Crear enlaces (hyperlinks) para cada hoja
        vLineaError = 70
        Call fun809_Crear_Enlace_Hoja(vHojaInventario, vFilaActual, 3, vNombreHoja)
        
        ' Paso 4.3: Determinar si la hoja es visible
        vLineaError = 80
        vEsVisible = (ThisWorkbook.Worksheets(vNombreHoja).Visible = xlSheetVisible)
        
        ' Paso 4.4: Aplicar formato segun visibilidad
        vLineaError = 90
        Call fun804_Aplicar_Formato_Inventario_Fila(vHojaInventario, vFilaActual, vEsVisible)
        
        ' Paso 4.5: Buscar fichero fuente en hoja "02_Log"
        vLineaError = 100
        vFicheroFuente = fun802_Buscar_Fichero_Fuente_En_Log(vNombreHoja)
        vHojaInventario.Cells(vFilaActual, 5).Value = vFicheroFuente
        
        vFilaActual = vFilaActual + 1
    Next vContadorHojas
    
    ' Paso 5: Ordenar el listado alfabeticamente (compatible Excel 97)
    vLineaError = 110
    If vFilaActual > 3 Then
        Set vRangoOrdenar = vHojaInventario.Range("B3:E" & (vFilaActual - 1))
        vRangoOrdenar.Sort Key1:=vHojaInventario.Range("B3"), Order1:=xlAscending, Header:=xlNo
    End If
    
    ' Paso 6: Asegurar visibilidad de hoja "01_Inventario"
    vLineaError = 120
    vHojaInventario.Visible = xlSheetVisible
    
    ' Ajustar columnas automaticamente
    vLineaError = 125
    vHojaInventario.Columns("B:E").AutoFit
    
    F012_Inventariar_Hojas = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F012_Inventariar_Hojas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description
    
    MsgBox vMensajeError, vbCritical, "Error F012_Inventariar_Hojas"
    
    If vResultado = 0 Then
        vResultado = 3999 ' Error no especificado en inventario
    End If
    
    F012_Inventariar_Hojas = vResultado
    
End Function

' =============================================================================
' FUNCION AUXILIAR: fun801_Configurar_Visibilidad_Hojas_Especificas
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Configura la visibilidad de hojas especificas segun criterios
' PARAMETROS: vNombreHoja (String)
' =============================================================================
Public Sub fun801_Configurar_Visibilidad_Hojas_Especificas(vNombreHoja As String)
    
    On Error GoTo ErrorHandler
    
    Dim vHojasVisibles(1 To 4) As String
    Dim vHojasOcultas(1 To 2) As String
    Dim i As Integer
    Dim vHoja As Worksheet
    
    ' Definir hojas que deben estar visibles
    vHojasVisibles(1) = "00_Ejecutar_Procesos"
    vHojasVisibles(2) = "01_Inventario"
    vHojasVisibles(3) = "02_Log"
    vHojasVisibles(4) = "09_Report_PL"
    
    ' Definir hojas que deben estar ocultas
    vHojasOcultas(1) = "05_Username"
    vHojasOcultas(2) = "06_Delimitadores_Originales"
    
    Set vHoja = ThisWorkbook.Worksheets(vNombreHoja)
    
    ' Verificar si debe estar visible
    For i = 1 To 4
        If StrComp(vNombreHoja, vHojasVisibles(i), vbTextCompare) = 0 Then
            vHoja.Visible = xlSheetVisible
            Exit Sub
        End If
    Next i
    
    ' Verificar si debe estar oculta
    For i = 1 To 2
        If StrComp(vNombreHoja, vHojasOcultas(i), vbTextCompare) = 0 Then
            vHoja.Visible = xlSheetHidden
            Exit Sub
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

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
    Set vHojaLog = ThisWorkbook.Worksheets("02_Log")
    If vHojaLog Is Nothing Then Exit Function
    
    ' Determinar texto a buscar basado en el nombre de la hoja
    If Left(vNombreHoja, 13) = "Import_Envio_" Then
        vBuscarTexto = "Import_" & Right(vNombreHoja, 15)
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

' =============================================================================
' FUNCION AUXILIAR: fun803_Aplicar_Formato_Inventario_Encabezados
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Aplica formato a los encabezados del inventario
' PARAMETROS: vHojaInventario (Worksheet)
' =============================================================================
Public Sub fun803_Aplicar_Formato_Inventario_Encabezados(vHojaInventario As Worksheet)
    
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

' =============================================================================
' FUNCION AUXILIAR: fun804_Aplicar_Formato_Inventario_Fila
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Aplica formato a una fila del inventario segun visibilidad
' PARAMETROS: vHojaInventario (Worksheet), vFila (Integer), vEsVisible (Boolean)
' =============================================================================
Public Sub fun804_Aplicar_Formato_Inventario_Fila(vHojaInventario As Worksheet, vFila As Integer, vEsVisible As Boolean)
    
    On Error GoTo ErrorHandler
    
    Dim vRangoFila As Range
    
    ' Definir rango de la fila (columnas 2 a 4)
    Set vRangoFila = vHojaInventario.Range("B" & vFila & ":D" & vFila)
    
    If vEsVisible Then
        ' Fila visible: sin color de fondo
        vRangoFila.Interior.ColorIndex = xlNone
        vHojaInventario.Cells(vFila, 4).Value = ">> visible <<"
    Else
        ' Fila oculta: fondo gris medio
        vRangoFila.Interior.Color = RGB(128, 128, 128)
        vHojaInventario.Cells(vFila, 4).Value = "OCULTA"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

' =============================================================================
' FUNCION AUXILIAR: fun805_Es_Hoja_Protegida
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Verifica si una hoja esta en la lista de hojas protegidas
' PARAMETROS: vNombreHoja (String)
' RETORNO: Boolean (True=protegida, False=no protegida)
' =============================================================================
Public Function fun805_Es_Hoja_Protegida(vNombreHoja As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim vHojasProtegidas(1 To 5) As String
    Dim i As Integer
    
    ' Lista de hojas protegidas
    vHojasProtegidas(1) = "00_Ejecutar_Procesos"
    vHojasProtegidas(2) = "01_Inventario"
    vHojasProtegidas(3) = "05_Username"
    vHojasProtegidas(4) = "06_Delimitadores_Originales"
    vHojasProtegidas(5) = "09_Report_PL"
    
    fun805_Es_Hoja_Protegida = False
    
    For i = 1 To 5
        If StrComp(vNombreHoja, vHojasProtegidas(i), vbTextCompare) = 0 Then
            fun805_Es_Hoja_Protegida = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    fun805_Es_Hoja_Protegida = False
    
End Function

' =============================================================================
' SUB AUXILIAR: fun806_Eliminar_Hoja_Segura
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Elimina una hoja de forma segura con control de errores
' PARAMETROS: vNombreHoja (String)
' =============================================================================
Public Sub fun806_Eliminar_Hoja_Segura(vNombreHoja As String)
    
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

' =============================================================================
' SUB AUXILIAR: fun807_Gestionar_Hoja_Envio
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Gestiona visibilidad de hojas Import_Envio_ segun antiguedad
' PARAMETROS: vNombreHoja (String), vHojasEnvio (Array), vNumHojasEnvio (Integer)
' =============================================================================
Public Sub fun807_Gestionar_Hoja_Envio(vNombreHoja As String, vHojasEnvio() As String, vNumHojasEnvio As Integer)
    
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
    If vNumHojasEnvio > vNum_Hojas_Hcas_Visibles_Envio And vPosicion > vNum_Hojas_Hcas_Visibles_Envio Then
        vHoja.Visible = xlSheetHidden
    Else
        vHoja.Visible = xlSheetVisible
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

' =============================================================================
' FUNCION AUXILIAR: fun808_Obtener_Hoja_Inventario
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Obtiene referencia a la hoja de inventario
' RETORNO: Worksheet (objeto hoja o Nothing si error)
' =============================================================================
Public Function fun808_Obtener_Hoja_Inventario() As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set fun808_Obtener_Hoja_Inventario = ThisWorkbook.Worksheets("01_Inventario")
    Exit Function
    
ErrorHandler:
    Set fun808_Obtener_Hoja_Inventario = Nothing
    
End Function

' =============================================================================
' SUB AUXILIAR: fun809_Crear_Enlace_Hoja
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Crea un hyperlink a una hoja especifica (compatible Excel 97)
' PARAMETROS: vHojaDestino, vFila, vColumna, vNombreHoja
' =============================================================================
Public Sub fun809_Crear_Enlace_Hoja(vHojaDestino As Worksheet, vFila As Integer, vColumna As Integer, vNombreHoja As String)
    
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

' =============================================================================
' SUB PRINCIPAL: Ejecutar_Procesos_Inicio_Libro
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Sub principal que ejecuta todos los procesos al abrir el libro
' PARAMETROS: Ninguno
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Sub Ejecutar_Procesos_Inicio_Libro()

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Ejecutar F010_Abrir_Hoja_Inicial
    ' 2. Ejecutar F011_Limpieza_Hojas_Historicas
    ' 3. Ejecutar F012_Inventariar_Hojas
    ' 4. Evaluar resultados y manejar errores

    On Error GoTo ErrorHandler
    
    Dim vResultado1 As Integer
    Dim vResultado2 As Integer
    Dim vResultado3 As Integer
    Dim vLineaError As Integer
    
    vLineaError = 10
    
    ' Paso 1: Ejecutar F010_Abrir_Hoja_Inicial
    vLineaError = 20
    vResultado1 = F010_Abrir_Hoja_Inicial()
    
    ' Paso 2: Ejecutar F011_Limpieza_Hojas_Historicas
    vLineaError = 30
    vResultado2 = F011_Limpieza_Hojas_Historicas()
    
    ' Paso 3: Ejecutar F012_Inventariar_Hojas
    vLineaError = 40
    vResultado3 = F012_Inventariar_Hojas()
    
    ' Paso 4: Evaluar resultados (opcional mostrar resumen)
    ' If vResultado1 <> 0 Or vResultado2 <> 0 Or vResultado3 <> 0 Then
    '     MsgBox "Procesos completados con algunos errores:" & vbCrLf & _
    '            "Navegacion: " & vResultado1 & vbCrLf & _
    '            "Limpieza: " & vResultado2 & vbCrLf & _
    '            "Inventario: " & vResultado3, vbInformation, "Resumen Procesos"
    ' End If
    
    Exit Sub
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en Ejecutar_Procesos_Inicio_Libro" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description
    
    MsgBox vMensajeError, vbCritical, "Error Principal"
    
End Sub

