Attribute VB_Name = "Modulo_121_FUNC_Auxiliares"
Option Explicit

'******************************************************************************
' M�dulo: Fun_Utils_Import
' Fecha y Hora de Creaci�n: 2025-05-26 05:33:25 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripci�n:
' Funciones auxiliares para el proceso de importaci�n de archivos
'******************************************************************************

'------------------------------------------------------------------------------
' Funci�n: fun807_ValidateSheetNames
' Descripci�n: Valida que existan las hojas especificadas
'------------------------------------------------------------------------------
Public Function fun807_ValidateSheetNames(ByVal strHojaImportacion As String, _
                                        ByVal strHojaWorking As String, _
                                        ByVal strHojaEnvio As String) As Boolean
    On Error GoTo GestorErrores
    
    fun807_ValidateSheetNames = False
    
    ' Validar existencia de hojas
    If Not fun802_SheetExists(strHojaImportacion) Then
        MsgBox "No existe la hoja: " & strHojaImportacion, vbExclamation
        Exit Function
    End If
    
    If Not fun802_SheetExists(strHojaWorking) Then
        MsgBox "No existe la hoja: " & strHojaWorking, vbExclamation
        Exit Function
    End If
    
    If Not fun802_SheetExists(strHojaEnvio) Then
        MsgBox "No existe la hoja: " & strHojaEnvio, vbExclamation
        Exit Function
    End If
    
    fun807_ValidateSheetNames = True
    Exit Function
    
GestorErrores:
    fun807_ValidateSheetNames = False
End Function

'------------------------------------------------------------------------------
' Funci�n: fun808_ClearSheet
' Descripci�n: Limpia el contenido de una hoja espec�fica
'------------------------------------------------------------------------------
Public Function fun808_ClearSheet(ByVal strSheetName As String) As Boolean
    On Error GoTo GestorErrores
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(strSheetName)
    
    ' Guardar configuraci�n actual
    Dim blnScreenUpdating As Boolean
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Limpiar contenido manteniendo formatos
    ws.UsedRange.ClearContents
    
    fun808_ClearSheet = True
    
CleanExit:
    Application.ScreenUpdating = blnScreenUpdating
    Exit Function
    
GestorErrores:
    fun808_ClearSheet = False
    Resume CleanExit
End Function

'------------------------------------------------------------------------------
' Funci�n: fun809_GetImportFilePath
' Descripci�n: Muestra di�logo para seleccionar archivo a importar
'------------------------------------------------------------------------------
Public Function fun809_GetImportFilePath(ByVal strPrompt As String) As String
    On Error GoTo GestorErrores
    
    Dim fd As Object ' FileDialog
    Dim strPath As String
    
    ' Crear di�logo de archivo
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = strPrompt
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt;*.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then ' Si se seleccion� un archivo
            strPath = .SelectedItems(1)
            
            ' Validar el archivo seleccionado
            If Not fun812_ValidateImportFile(strPath) Then
                MsgBox "El archivo seleccionado no es v�lido.", vbExclamation
                strPath = ""
            End If
        End If
    End With
    
    fun809_GetImportFilePath = strPath
    Exit Function
    
GestorErrores:
    fun809_GetImportFilePath = ""
End Function

'------------------------------------------------------------------------------
' Funci�n: fun810_DetectDataRange
' Descripci�n: Detecta el rango de datos en la hoja de importaci�n
'------------------------------------------------------------------------------
Public Function fun810_DetectDataRange(ByRef ws As Worksheet, _
                                     ByRef lngLineaInicial As Long, _
                                     ByRef lngLineaFinal As Long) As Boolean
    On Error GoTo GestorErrores
    
    Dim rngDatos As Range
    
    ' Obtener rango usado en la columna inicial
    Set rngDatos = ws.Range(gstrColumnaInicial_Importacion & ":" & _
                           gstrColumnaInicial_Importacion).SpecialCells(xlCellTypeConstants)
    
    ' Determinar primera y �ltima fila con datos
    lngLineaInicial = rngDatos.Row
    lngLineaFinal = rngDatos.Row + rngDatos.Rows.Count - 1
    
    fun810_DetectDataRange = True
    Exit Function
    
GestorErrores:
    fun810_DetectDataRange = False
End Function

'------------------------------------------------------------------------------
' Funci�n: fun811_FormatColumns
' Descripci�n: Aplica formato espec�fico a las columnas
'------------------------------------------------------------------------------
Public Function fun811_FormatColumns(ByRef ws As Worksheet) As Boolean
    On Error GoTo GestorErrores
    
    Dim lngColInicio As Long
    Dim lngCol As Long
    Dim rngFormato As Range
    
    ' Convertir letra de columna a n�mero
    lngColInicio = Range(gstrColumnaInicial_Importacion & "1").Column
    
    ' Formato texto para columnas 1-10
    For lngCol = lngColInicio To lngColInicio + 10
        Set rngFormato = ws.Range(ws.Cells(glngLineaInicial_HojaImportacion, lngCol), _
                                ws.Cells(glngLineaFinal_HojaImportacion, lngCol))
        rngFormato.NumberFormat = "@"
    Next lngCol
    
    ' Formato num�rico para columnas 11-22
    For lngCol = lngColInicio + 11 To lngColInicio + 22
        Set rngFormato = ws.Range(ws.Cells(glngLineaInicial_HojaImportacion, lngCol), _
                                ws.Cells(glngLineaFinal_HojaImportacion, lngCol))
        rngFormato.NumberFormat = "#,##0.00"
    Next lngCol
    
    fun811_FormatColumns = True
    Exit Function
    
GestorErrores:
    fun811_FormatColumns = False
End Function

'------------------------------------------------------------------------------
' Funci�n: fun812_ValidateImportFile
' Descripci�n: Valida el archivo seleccionado para importaci�n
'------------------------------------------------------------------------------
Public Function fun812_ValidateImportFile(ByVal strFilePath As String) As Boolean
    On Error GoTo GestorErrores
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fun812_ValidateImportFile = False
    
    ' Validar que el archivo existe
    If Not fso.FileExists(strFilePath) Then
        MsgBox "El archivo no existe.", vbExclamation
        Exit Function
    End If
    
    ' Validar extensi�n
    Select Case LCase(fso.GetExtensionName(strFilePath))
        Case "txt", "csv"
            ' Extensiones v�lidas
            fun812_ValidateImportFile = True
        Case Else
            MsgBox "El archivo debe ser .txt o .csv", vbExclamation
    End Select
    
    Exit Function
    
GestorErrores:
    fun812_ValidateImportFile = False
End Function


