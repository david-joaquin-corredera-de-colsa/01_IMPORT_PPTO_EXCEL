Attribute VB_Name = "Modulo_501_SUB_principal_02"
Option Explicit
'resultado = Limpiar_Otra_Informacion()

Public Sub M004_Limpiar_Otra_Informacion()

    ' *********************************
    ' Variables para control de Errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    Dim blnInventarioActualizado As Boolean
    ' *********************************
    
    '--------------------------------------------------------------------------
    ' 0. Inicializacion
    '--------------------------------------------------------------------------
    strFuncion = "M004_Limpiar_Otra_Informacion" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Eliminar_Hojas_Antiguas"
    lngLineaError = 400000
    
    On Error GoTo ErrorHandler

    '--------------------------------------------------------------------------
    ' 1. Configuración inicial del entorno - Optimizacion
    '--------------------------------------------------------------------------
    lngLineaError = 400001
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Inicializar variables globales
    Call InitializeGlobalVariables
    
    fun801_LogMessage "Iniciando limpieza del log ..."

    '--------------------------------------------------------------------------
    ' 2. Eliminamos hojas no deseadas
    '--------------------------------------------------------------------------
    lngLineaError = 400002
    fun801_LogMessage "Limpiando otra informacion ... " & vbCrLf & "en el libro Excel de carga de Presupuesto ..."
       
    blnResult = Limpiar_Otra_Informacion()
    
    If Not blnResult Then
        MsgBox "Error al limpiar el log"
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'limpiar' otra informacion"
    Else
        MsgBox "Otra información " & vbCrLf & "se limpió con éxito"
    End If

    GoTo CleanExit ' Asegura que se ejecuta la limpieza aunque no haya error
    
CleanExit:

    '--------------------------------------------------------------------------
    ' 3. Ordenamos las hojas que quedan en el libro
    '--------------------------------------------------------------------------
    lngLineaError = 400003
    fun801_LogMessage "Ordenando las hojas del libro ..."
    
    blnResult = Ordenar_Hojas()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'ordenar' las hojas"
    End If


    '--------------------------------------------------------------------------
    ' 4. Comprobamos si el inventario esta actualizado
    '--------------------------------------------------------------------------
    lngLineaError = 400004
    fun801_LogMessage "Comprobando si el inventario esta actualizado ..."
    
    
    blnInventarioActualizado = Inventario_Actualizado_Si_No()
    
    If blnInventarioActualizado Then
        
        fun801_LogMessage "El inventario de hojas estaba actualizado (no será necesario actualizarlo ahora mismo)"
        
    Else
        
        fun801_LogMessage "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        'MsgBox "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        '--------------------------------------------------------------------------
        ' 5. Limpiamos las hojas "tecnicas" historicas
        '   borramos las que no son "Import_Envio_"
        '   dejamos visibles solo las ultimas X (5) hojas de "Import_Envio_"
        '   más recientes
        '--------------------------------------------------------------------------
        lngLineaError = 400005
        fun801_LogMessage "Limpiando las hojas tecnicas históricas ..."
        
        blnResult = F011_Limpieza_Hojas_Historicas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente la limpeza de hojas tecnicas históricas"
        End If
            
        '--------------------------------------------------------------------------
        ' 6. Actualizamos la hoja 01_Inventario
        '   Listando todas las hojas existentes en el libro
        '--------------------------------------------------------------------------
        lngLineaError = 400006
        fun801_LogMessage "Inventariando las hojas del libro ..."
        
        blnResult = F012_Inventariar_Hojas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente el inventario de hojas"
        End If
        'MsgBox "Acabamos de actualizar el inventario de hojas"
    End If

    '--------------------------------------------------------------------------
    ' 7. Abrir la hoja Principal
    '   (y desde alli ya se lanzan otras tareas iniciales
    '   propias de la apertura del libro)
    '--------------------------------------------------------------------------
    lngLineaError = 400007
    fun801_LogMessage "Regresando a la hoja inicial ..."
    
    blnResult = F010_Abrir_Hoja_Inicial()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se consiguió regresar correctamente a la hoja inicial"
    End If
    
    '--------------------------------------------------------------------------
    ' 9. Restauración del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ThisWorkbook.Save
    Exit Sub

ErrorHandler:
    ' Construcción del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Origen: " & Err.Source & vbCrLf & _
                      "Descripción: " & Err.Description
    
    ' Registro del error
    fun801_LogMessage strMensajeError, True
    
    ' Restaurar las opciones de optimizacion de entorno antes de salir
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    ' Mostrar mensaje al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    
    ' Asegurar que se restaura la configuración de Excel
    Resume CleanExit

End Sub
Public Sub M003_Limpiar_Log()

    ' *********************************
    ' Variables para control de Errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    Dim blnInventarioActualizado As Boolean
    ' *********************************
    
    '--------------------------------------------------------------------------
    ' 0. Inicializacion
    '--------------------------------------------------------------------------
    strFuncion = "M003_Limpiar_Log" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Eliminar_Hojas_Antiguas"
    lngLineaError = 500000
    
    On Error GoTo ErrorHandler

    '--------------------------------------------------------------------------
    ' 1. Configuración inicial del entorno - Optimizacion
    '--------------------------------------------------------------------------
    lngLineaError = 500001
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Inicializar variables globales
    Call InitializeGlobalVariables
    
    fun801_LogMessage "Iniciando limpieza del log ..."

    '--------------------------------------------------------------------------
    ' 2. Eliminamos hojas no deseadas
    '--------------------------------------------------------------------------
    lngLineaError = 500002
    fun801_LogMessage "Limpiando el log ... " & vbCrLf & "para eliminar entradas más antiguas ..."
       
    blnResult = Limpiar_Log()
    
    If Not blnResult Then
        MsgBox "Error al limpiar el log"
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'limpiar' el log"
    Else
        MsgBox "Log limpiado exitosamente"
    End If



    GoTo CleanExit ' Asegura que se ejecuta la limpieza aunque no haya error
    
CleanExit:

    '--------------------------------------------------------------------------
    ' 3. Ordenamos las hojas que quedan en el libro
    '--------------------------------------------------------------------------
    lngLineaError = 500003
    fun801_LogMessage "Ordenando las hojas del libro ..."
    
    blnResult = Ordenar_Hojas()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'ordenar' las hojas"
    End If


    '--------------------------------------------------------------------------
    ' 4. Comprobamos si el inventario esta actualizado
    '--------------------------------------------------------------------------
    lngLineaError = 500004
    fun801_LogMessage "Comprobando si el inventario esta actualizado ..."
    
    
    blnInventarioActualizado = Inventario_Actualizado_Si_No()
    
    If blnInventarioActualizado Then
        
        fun801_LogMessage "El inventario de hojas estaba actualizado (no será necesario actualizarlo ahora mismo)"
        
    Else
        
        fun801_LogMessage "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        'MsgBox "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        '--------------------------------------------------------------------------
        ' 5. Limpiamos las hojas "tecnicas" historicas
        '   borramos las que no son "Import_Envio_"
        '   dejamos visibles solo las ultimas X (5) hojas de "Import_Envio_"
        '   más recientes
        '--------------------------------------------------------------------------
        lngLineaError = 500005
        fun801_LogMessage "Limpiando las hojas tecnicas históricas ..."
        
        blnResult = F011_Limpieza_Hojas_Historicas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente la limpeza de hojas tecnicas históricas"
        End If
            
        '--------------------------------------------------------------------------
        ' 6. Actualizamos la hoja 01_Inventario
        '   Listando todas las hojas existentes en el libro
        '--------------------------------------------------------------------------
        lngLineaError = 500006
        fun801_LogMessage "Inventariando las hojas del libro ..."
        
        blnResult = F012_Inventariar_Hojas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente el inventario de hojas"
        End If
        'MsgBox "Acabamos de actualizar el inventario de hojas"
    End If

    '--------------------------------------------------------------------------
    ' 7. Abrir la hoja Principal
    '   (y desde alli ya se lanzan otras tareas iniciales
    '   propias de la apertura del libro)
    '--------------------------------------------------------------------------
    lngLineaError = 500007
    fun801_LogMessage "Regresando a la hoja inicial ..."
    
    blnResult = F010_Abrir_Hoja_Inicial()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se consiguió regresar correctamente a la hoja inicial"
    End If

    '--------------------------------------------------------------------------
    ' 9. Restauración del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ThisWorkbook.Save
    Exit Sub

ErrorHandler:
    ' Construcción del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Origen: " & Err.Source & vbCrLf & _
                      "Descripción: " & Err.Description
    
    ' Registro del error
    fun801_LogMessage strMensajeError, True
    
    ' Restaurar las opciones de optimizacion de entorno antes de salir
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    ' Mostrar mensaje al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    
    ' Asegurar que se restaura la configuración de Excel
    Resume CleanExit

End Sub
Public Sub M002_Eliminar_Hojas_Antiguas()

    ' *********************************
    ' Variables para control de Errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim blnResult As Boolean
    Dim blnInventarioActualizado As Boolean
    ' *********************************
    
    '--------------------------------------------------------------------------
    ' 0. Inicializacion
    '--------------------------------------------------------------------------
    strFuncion = "M002_Eliminar_Hojas_Antiguas" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Eliminar_Hojas_Antiguas"
    lngLineaError = 9980000
    
    On Error GoTo ErrorHandler

    '--------------------------------------------------------------------------
    ' 1. Configuración inicial del entorno - Optimizacion
    '--------------------------------------------------------------------------
    lngLineaError = 9980001
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Inicializar variables globales
    Call InitializeGlobalVariables
    
    fun801_LogMessage "Iniciando eliminacion de hojas antiguas ..."

    '--------------------------------------------------------------------------
    ' 2. Eliminamos hojas no deseadas
    '--------------------------------------------------------------------------
    lngLineaError = 9980002
    fun801_LogMessage "Eliminando hojas no deseadas ..."
    
    
    Dim vNumHojasTarget As Integer
    Dim vBorrarO_00 As Boolean
    Dim vBorrarO_I As Boolean
    Dim vBorrarO As Boolean
        
    
    vNumHojasTarget = InputBox("Numero de hojas 'Import_Envio_'" & vbCrLf & "que desea dejar", "Numero de hojas 'Import_Envio_'" & vbCrLf & "que desea dejar", CONST_NUM_HOJAS_HCAS_IMPORT_TARGET)
        'Donde normalmente CONST_NUM_HOJAS_HCAS_IMPORT_TARGET tiene un valor de 5
    vBorrarO_00 = CONST_BORRAR_OTRAS_HOJAS_PREFIJO_00
    vBorrarO_I = CONST_BORRAR_OTRAS_HOJAS_PREFIJO_IMPORT
    vBorrarO = CONST_BORRAR_HOJAS_SIN_PREFIJOS_00_IMPORT
    
    blnResult = Eliminar_Hojas_NoDeseadas(vNumHojasTarget, vBorrarO_00, vBorrarO_I, vBorrarO)
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'eliminar' las hojas no deseadas"
    Else
        MsgBox "Se eliminaron con exito" & vbCrLf & "las hojas antigüas" & vbCrLf & "y hojas no deseadas"
    End If


    GoTo CleanExit ' Asegura que se ejecuta la limpieza aunque no haya error
    
CleanExit:

    '--------------------------------------------------------------------------
    ' 3. Ordenamos las hojas que quedan en el libro
    '--------------------------------------------------------------------------
    lngLineaError = 9980003
    fun801_LogMessage "Ordenando las hojas del libro ..."
    
    blnResult = Ordenar_Hojas()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se completó correctamente la tarea de 'ordenar' las hojas"
    End If


    '--------------------------------------------------------------------------
    ' 4. Comprobamos si el inventario esta actualizado
    '--------------------------------------------------------------------------
    lngLineaError = 9980004
    fun801_LogMessage "Comprobando si el inventario esta actualizado ..."
    
    
    blnInventarioActualizado = Inventario_Actualizado_Si_No()
    
    If blnInventarioActualizado Then
        
        fun801_LogMessage "El inventario de hojas estaba actualizado (no será necesario actualizarlo ahora mismo)"
        
    Else
        
        fun801_LogMessage "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        'MsgBox "El inventario de hojas NO estaba actualizado y procederemos a actualizarlo ahora mismo"
        '--------------------------------------------------------------------------
        ' 5. Limpiamos las hojas "tecnicas" historicas
        '   borramos las que no son "Import_Envio_"
        '   dejamos visibles solo las ultimas X (5) hojas de "Import_Envio_"
        '   más recientes
        '--------------------------------------------------------------------------
        lngLineaError = 9980005
        fun801_LogMessage "Limpiando las hojas tecnicas históricas ..."
        
        blnResult = F011_Limpieza_Hojas_Historicas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente la limpeza de hojas tecnicas históricas"
        End If
            
        '--------------------------------------------------------------------------
        ' 6. Actualizamos la hoja 01_Inventario
        '   Listando todas las hojas existentes en el libro
        '--------------------------------------------------------------------------
        lngLineaError = 9980006
        fun801_LogMessage "Inventariando las hojas del libro ..."
        
        blnResult = F012_Inventariar_Hojas()
        
        If Not blnResult Then
            Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
                "No se ejecutó correctamente el inventario de hojas"
        End If
        'MsgBox "Acabamos de actualizar el inventario de hojas"
    End If

    '--------------------------------------------------------------------------
    ' 7. Abrir la hoja Principal
    '   (y desde alli ya se lanzan otras tareas iniciales
    '   propias de la apertura del libro)
    '--------------------------------------------------------------------------
    lngLineaError = 9980007
    fun801_LogMessage "Regresando a la hoja inicial ..."
    
    blnResult = F010_Abrir_Hoja_Inicial()
    
    If Not blnResult Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "No se consiguió regresar correctamente a la hoja inicial"
    End If
    
    
    '--------------------------------------------------------------------------
    ' 9. Restauración del entorno
    '--------------------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ThisWorkbook.Save
    Exit Sub

ErrorHandler:
    ' Construcción del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Origen: " & Err.Source & vbCrLf & _
                      "Descripción: " & Err.Description
    
    ' Registro del error
    fun801_LogMessage strMensajeError, True
    
    ' Restaurar las opciones de optimizacion de entorno antes de salir
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    ' Mostrar mensaje al usuario
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    
    ' Asegurar que se restaura la configuración de Excel
    Resume CleanExit

End Sub


