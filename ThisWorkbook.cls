VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' CODIGO PARA MODULO ThisWorkbook
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Evento Workbook_Open que ejecuta todos los procesos de inicio
' AUTOR: david-joaquin-corredera-de-colsa
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================

Private Sub Workbook_Open()
    
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Ejecutar macro de inicializacion completa del libro
    ' 2. Control de errores centralizado para el evento de apertura
    
    On Error GoTo ErrorHandler
    
    ' Paso 1: Ejecutar todos los procesos de inicio del libro
    ' Esta llamada ejecuta: F010_Abrir_Hoja_Inicial, F011_Limpieza_Hojas_Historicas, F012_Inventariar_Hojas
    Call Ejecutar_Procesos_Inicio_Libro
    
    Exit Sub
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en Workbook_Open" & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & vbCrLf & _
                   "El libro se abrira sin ejecutar los procesos automaticos."
    
    MsgBox vMensajeError, vbExclamation, "Error al Abrir Libro"
    
End Sub

