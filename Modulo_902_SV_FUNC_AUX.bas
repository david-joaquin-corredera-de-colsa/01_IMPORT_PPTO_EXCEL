Attribute VB_Name = "Modulo_902_SV_FUNC_AUX"
Option Explicit


Public Function SmartView_Options_MemberOptions_Indent_None(vNombreHoja As Variant) As Long
    SmartView_Options_MemberOptions_Indent_None = HypSetSheetOption(vNombreHoja, CONST_INDENT_SETTING, CONST_INDENT_NONE)
    If SmartView_Options_MemberOptions_Indent_None = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Member Options > Indent = None"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Member Options > Indent = None." & vbCrLf & "Error Number = " & SmartView_Options_MemberOptions_Indent_None
    End If
End Function

Public Function SmartView_Options_DataOptions_SUPPRESS_Missing(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_SUPPRESS_Missing = HypSetSheetOption(vNombreHoja, CONST_SUPPRESS_MISSING_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_SUPPRESS_Missing = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Missing = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Data Options > SUPPRESS Missing = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_SUPPRESS_Missing
    End If
    
End Function
Public Function SmartView_Options_DataOptions_SUPPRESS_Zero(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_SUPPRESS_Zero = HypSetSheetOption(vNombreHoja, CONST_SUPPRESS_ZERO_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_SUPPRESS_Zero = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Zero = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Zero = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_SUPPRESS_Zero
    End If
    
End Function

Public Function SmartView_Options_DataOptions_SUPPRESS_Repeated(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_SUPPRESS_Repeated = HypSetSheetOption(vNombreHoja, CONST_ENABLE_REPEATED_MEMBERS_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_SUPPRESS_Repeated = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Repeated = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Repeated = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_SUPPRESS_Repeated
    End If
End Function
Public Function SmartView_Options_DataOptions_SUPPRESS_Invalid(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_SUPPRESS_Invalid = HypSetSheetOption(vNombreHoja, CONST_ENABLE_INVALID_MEMBERS_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_SUPPRESS_Invalid = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Invalid = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS Invalid = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_SUPPRESS_Invalid
    End If
    
End Function
Public Function SmartView_Options_DataOptions_SUPPRESS_NoAccess(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_SUPPRESS_NoAccess = HypSetSheetOption(vNombreHoja, CONST_ENABLE_NOACCESS_MEMBERS_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_SUPPRESS_NoAccess = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS NoAccess = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > SUPPRESS NoAccess = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_SUPPRESS_NoAccess
    End If
End Function

Public Function SmartView_Options_MemberOptions_PreserveFormulasComments(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_MemberOptions_PreserveFormulasComments = HypSetOption(HSV_PRESERVE_FORMULA_COMMENT, vTrueFalse, vNombreHoja)
    If SmartView_Options_MemberOptions_PreserveFormulasComments = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Member Options > Preserve Formulas and Comments = True"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Member Options > Preserve Formulas and Comments = True." & vbCrLf & "Error Number = " & SmartView_Options_MemberOptions_PreserveFormulasComments
    End If
End Function

Public Function SmartView_Options_DataOptions_CellDisplay(vNombreHoja As Variant) As Long
    SmartView_Options_DataOptions_CellDisplay = HypSetSheetOption(vNombreHoja, CONST_CELL_DISPLAY_SETTING, CONST_CELL_DISPLAY_SHOW_DATA)
    If SmartView_Options_DataOptions_CellDisplay = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Cell Display = Data"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > Cell Display = Data." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_CellDisplay
    End If
End Function

Public Function SmartView_Options_MemberOptions_DisplayNameOnly(vNombreHoja As Variant) As Long
    SmartView_Options_MemberOptions_DisplayNameOnly = HypSetSheetOption(vNombreHoja, CONST_DISPLAY_MEMBER_NAME_SETTING, CONST_DISPLAY_NAME_ONLY)
    If SmartView_Options_MemberOptions_DisplayNameOnly = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Member Options > Member Display = Name Only"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Member Options > Member Display = Name Only." & vbCrLf & "Error Number = " & SmartView_Options_MemberOptions_DisplayNameOnly
    End If
End Function


