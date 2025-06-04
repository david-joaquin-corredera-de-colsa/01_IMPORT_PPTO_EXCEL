VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Username_Password 
   Caption         =   "Introduzca su Username y Password de HFM"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm_Username_Password.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Username_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton_Aceptar_Click()
    If Trim(TextBox_Username.Value) <> "" And Trim(TextBox_Password.Value) <> "" Then
        Me.Hide
    ElseIf Trim(TextBox_Username.Value) <> "" Then
        MsgBox "Debe introducir su Password"
    ElseIf Trim(TextBox_Password.Value) <> "" Then
        MsgBox "Debe introducir su Username"
    End If




End Sub

Private Sub CommandButton_Cancelar_Click()
    Me.Hide
End Sub
